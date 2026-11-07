from __future__ import annotations

"""Read-only native audit of RefreshAll and its nested VBA dialog sites.

No macros are executed and no workbook is saved. This tool identifies candidate
MsgBox/InputBox calls, rather than falsely claiming any of them ran at runtime.
"""

import argparse
from collections import defaultdict
from pathlib import Path
import re
import sys

START = re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)\s+([A-Za-z][A-Za-z_0-9]*)\b", re.I)
END = re.compile(r"^\s*End\s+(?:Sub|Function)\b", re.I)
DIALOG = re.compile(r"\b(?:MsgBox|InputBox|Application\.InputBox)\s*(?:\(|\b)", re.I)
CALL = re.compile(r"^\s*(?:Call\s+)?([A-Za-z][A-Za-z_0-9]*)\b", re.I)
REFRESH = (
    "RefreshTherapistDaily",
    "RefreshTherapistAttendanceAndConflicts",
    "BuildReplacementSuggestions_V4",
    "ColorInfectiousInMaster",
    "ColorReplacementMarks",
    "RestoreReturnedTherapists",
)


def extract_procedures(modules: dict[str, str]) -> dict[str, list[tuple[str, int, tuple[str, ...]]]]:
    result: dict[str, list[tuple[str, int, tuple[str, ...]]]] = defaultdict(list)
    for module, source in modules.items():
        lines = source.splitlines()
        i = 0
        while i < len(lines):
            found = START.match(lines[i])
            if not found:
                i += 1
                continue
            begin = i
            i += 1
            while i < len(lines) and not END.match(lines[i]):
                i += 1
            if i >= len(lines):
                raise ValueError(f"Unterminated VBA procedure in {module}:{begin + 1}")
            result[found.group(1).casefold()].append((module, begin + 1, tuple(lines[begin:i + 1])))
            i += 1
    return dict(result)


def dialog_audit(modules: dict[str, str]) -> tuple[list[str], list[str]]:
    procs = extract_procedures(modules)
    missing = []
    records = []
    queue = ["RefreshAll", *REFRESH]
    visited: set[str] = set()
    while queue:
        name = queue.pop(0)
        norm = name.casefold()
        if norm in visited:
            continue
        visited.add(norm)
        matches = procs.get(norm, [])
        if not matches:
            missing.append(name)
            continue
        if len(matches) != 1:
            raise ValueError(f"Ambiguous VBA procedure {name}: {len(matches)} definitions")
        module, start, lines = matches[0]
        records.append(f"PROCEDURE {name} {module}:{start}")
        for offset, line in enumerate(lines):
            code = line.split("'", 1)[0].strip()
            if DIALOG.search(code):
                records.append(f"DIALOG CANDIDATE {module}:{start + offset} {line.strip()}")
            if offset == 0 or not code or code.casefold().startswith(("rem ", "on error", "end ", "exit ")):
                continue
            match = CALL.match(code)
            if match:
                callee = match.group(1)
                if callee.casefold() in procs and callee.casefold() not in visited:
                    queue.append(callee)
    return records, missing


def audit(workbook_path: Path) -> int:
    if sys.platform != "win32" or not workbook_path.is_file() or workbook_path.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: Existing XLSM on Windows required")
        return 2
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 required")
        return 2
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=True)
        if not bool(wb.ReadOnly):
            raise RuntimeError("ReadOnly contract violated")
        modules = {}
        components = wb.VBProject.VBComponents
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            code = component.CodeModule
            if code.CountOfLines:
                modules[str(component.Name)] = str(code.Lines(1, code.CountOfLines))
        records, missing = dialog_audit(modules)
        print("H8 REFRESH DIALOG READ-ONLY AUDIT")
        print("Workbook:", wb.FullName)
        print("ReadOnly:", bool(wb.ReadOnly))
        for row in records:
            print(row)
        for name in missing:
            print("MISSING PROCEDURE:", name)
        print("Workbook unchanged: True")
        if missing:
            print("H8 REFRESH AUDIT INCOMPLETE")
            return 3
        print("H8 REFRESH DIALOG AUDIT COMPLETE; candidates are not confirmed runtime prompts")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only RefreshAll VBA call/dialog inspection")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return audit(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
