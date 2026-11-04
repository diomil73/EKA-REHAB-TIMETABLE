from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys


PROCEDURE = re.compile(
    r"^\\s*(?:(?:Public|Private|Friend|Static)\\s+)?"
    r"(?:Sub|Function|Property\\s+(?:Get|Let|Set))\\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)
WRITE_LINE = re.compile(
    r"\\.\\s*(?:Unprotect|Protect|Value2?|Formula|ClearContents|Clear|Delete|Insert|Add|"
    r"PasteSpecial|Copy|Locked|Visible|ScrollArea)\\b"
    r"|\\b(?:Worksheets|Sheets)\\s*\\(\\s*[\"'][^\"']+[\"']\\s*\\)"
    r"|\\b(?:ThisWorkbook|ActiveWorkbook)\\.(?:Save|SaveAs)\\b"
    r"|\\bDestination\\s*:",
    re.IGNORECASE,
)
MAX_MATCHES_PER_COMPONENT = 180
MAX_TOTAL_MATCHES = 2400


def _matches(lines: list[str]) -> list[tuple[int, str, str]]:
    current = "(module declarations)"
    found: list[tuple[int, str, str]] = []
    for number, line in enumerate(lines, start=1):
        match = PROCEDURE.match(line)
        if match:
            current = match.group(1)
        if WRITE_LINE.search(line):
            found.append((number, current, line.strip()))
    return found


def audit(path: Path) -> int:
    if sys.platform != "win32" or not path.is_file() or path.suffix.lower() != ".xlsm":
        print("SAFETY STOP: existing XLSM on Windows is required")
        return 2
    try:
        import win32com.client
    except ImportError:
        print("SAFETY STOP: pywin32 is required")
        return 2
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        if not bool(workbook.ReadOnly):
            print("SAFETY STOP: workbook did not open read-only")
            return 2
        components = workbook.VBProject.VBComponents
        count = int(components.Count)
        print("VBA WRITE-SURFACE READ-ONLY AUDIT")
        print("Workbook:", workbook.FullName)
        print("ReadOnly:", bool(workbook.ReadOnly))
        print("VBA component count:", count)
        total = 0
        truncated = False
        for index in range(1, count + 1):
            component = components(index)
            module = component.CodeModule
            line_count = int(module.CountOfLines)
            if line_count <= 0:
                continue
            raw = str(module.Lines(1, line_count))
            lines = raw.replace("\\r\\n", "\\n").replace("\\r", "\\n").split("\\n")
            matches = _matches(lines)
            if not matches:
                continue
            print(f"COMPONENT {component.Name}: type={int(component.Type)} lines={line_count} matches={len(matches)}")
            for line_number, procedure, source in matches[:MAX_MATCHES_PER_COMPONENT]:
                if total >= MAX_TOTAL_MATCHES:
                    truncated = True
                    break
                print(f"  {procedure}:{line_number}: {source}")
                total += 1
            if len(matches) > MAX_MATCHES_PER_COMPONENT:
                truncated = True
                print("  COMPONENT OUTPUT TRUNCATED")
            if truncated and total >= MAX_TOTAL_MATCHES:
                break
        print("Matched source lines:", total)
        print("Workbook unchanged: True")
        if truncated:
            print("VBA WRITE-SURFACE AUDIT INCOMPLETE: output cap reached")
            return 3
        print("VBA WRITE-SURFACE AUDIT COMPLETE")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only inventory of VBA worksheet write surfaces")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return audit(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
