from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import re
import sys


PROCEDURE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static)\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)
WRITE_LINE = re.compile(
    r"\.\s*(?:Unprotect|Protect|Value2?|Formula|NumberFormat|Font|Interior|Borders|"
    r"FormatConditions|Validation|ClearContents|Clear|Delete|Insert|Add|PasteSpecial|Copy|"
    r"Locked|Visible|Hidden|ScrollArea|EnableSelection|ColumnWidth|RowHeight|Name|OnAction)\b"
    r"|\b(?:Worksheets|Sheets)\s*\("
    r"|\b(?:ThisWorkbook|ActiveWorkbook)\.(?:Save|SaveAs)\b"
    r"|\bDestination\s*:",
    re.IGNORECASE,
)
MAX_MATCHES_PER_COMPONENT = 200
MAX_TOTAL_MATCHES = 3000


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


def _operation(line: str) -> tuple[str, str | None]:
    lowered = line.lower()
    sheet = re.search(r'\\b(?:Worksheets|Sheets)\\s*\\(\\s*"([^"]+)"', line, re.IGNORECASE)
    target = sheet.group(1) if sheet else None
    if re.search(r"\\.\\s*Unprotect\\b", line, re.IGNORECASE):
        return "protection_unprotect", target
    if re.search(r"\\.\\s*Protect\\b", line, re.IGNORECASE):
        return "protection_protect", target
    if re.search(r"\\b(?:ThisWorkbook|ActiveWorkbook)\\.(?:Save|SaveAs)\\b", line, re.IGNORECASE):
        return "workbook_save", target
    if re.search(r"\\.\\s*(?:ClearContents|Clear|Delete|Insert|Add|PasteSpecial|Copy)\\b", line, re.IGNORECASE):
        return "range_or_structure_mutation", target
    if re.search(r"\\.\\s*(?:Value2?|Formula)\\b", line, re.IGNORECASE):
        return "cell_value_or_formula_write", target
    if re.search(r"\\.\\s*(?:Visible|Hidden|ScrollArea|EnableSelection|OnAction)\\b", line, re.IGNORECASE):
        return "visibility_navigation_or_action_binding", target
    if target:
        return "worksheet_reference", target
    return "format_or_other_write", target


def _fingerprint(path: Path) -> tuple[str, int]:
    digest = sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def audit(path: Path) -> int:
    if sys.platform != "win32" or not path.is_file() or path.suffix.lower() != ".xlsm":
        print("SAFETY STOP: existing XLSM on Windows is required")
        return 2
    before_hash, before_size = _fingerprint(path)
    try:
        import win32com.client
    except ImportError:
        print("SAFETY STOP: pywin32 is required")
        return 2
    excel = None
    workbook = None
    status = 0
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        # msoAutomationSecurityForceDisable: prevent workbook macros from running.
        excel.AutomationSecurity = 3
        workbook = excel.Workbooks.Open(
            str(path.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
        )
        if not bool(workbook.ReadOnly):
            print("SAFETY STOP: workbook did not open read-only")
            status = 2
        else:
            components = workbook.VBProject.VBComponents
            count = int(components.Count)
            print("VBA WRITE-SURFACE READ-ONLY AUDIT")
            print("Workbook:", path.name)
            print("ReadOnly: True")
            print("Macros disabled before open: YES")
            print("Excel events disabled: YES")
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
                print(
                    f"COMPONENT {component.Name}: type={int(component.Type)} "
                    f"lines={line_count} matched_operations={len(matches)}"
                )
                for line_number, procedure, source in matches[:MAX_MATCHES_PER_COMPONENT]:
                    if total >= MAX_TOTAL_MATCHES:
                        truncated = True
                        break
                    operation, target = _operation(source)
                    suffix = f" target_sheet={target}" if target else ""
                    print(
                        f"  procedure={procedure} line={line_number} "
                        f"operation={operation}{suffix}"
                    )
                    total += 1
                if len(matches) > MAX_MATCHES_PER_COMPONENT:
                    truncated = True
                    print("  COMPONENT OUTPUT TRUNCATED")
                if truncated and total >= MAX_TOTAL_MATCHES:
                    break
            print("Matched operations:", total)
            if truncated:
                print("VBA WRITE-SURFACE AUDIT INCOMPLETE: output cap reached")
                status = 3
            else:
                print("VBA WRITE-SURFACE AUDIT COMPLETE")
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        status = 2
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                status = max(status, 2)
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                status = max(status, 2)
    try:
        after_hash, after_size = _fingerprint(path)
        unchanged = before_hash == after_hash and before_size == after_size
        print("SOURCE HASH UNCHANGED:", "YES" if unchanged else "NO")
        print("NO WORKBOOK WAS MODIFIED" if unchanged else "SAFETY STOP: workbook fingerprint changed")
        if not unchanged:
            status = max(status, 2)
    except Exception as exc:
        print(f"SAFETY STOP: unable to verify source fingerprint ({type(exc).__name__})")
        status = max(status, 2)
    return status

def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only inventory of VBA worksheet write surfaces")
    parser.add_argument("workbooks", type=Path, nargs="+")
    args = parser.parse_args()
    overall = 0
    seen: set[str] = set()
    for workbook in args.workbooks:
        if not workbook.is_file():
            print(f"SAFETY STOP: workbook not found: {workbook}")
            overall = max(overall, 2)
            continue
        digest, _ = _fingerprint(workbook)
        if digest in seen:
            print(f"BYTE-IDENTICAL DUPLICATE SKIPPED: {workbook.name}")
            continue
        seen.add(digest)
        result = audit(workbook)
        overall = max(overall, result)
    return overall


if __name__ == "__main__":
    raise SystemExit(main())
