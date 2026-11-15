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
    r"|\bDestination\s*:"
    r"|\bKill\s+|\b(?:FileCopy|MkDir|RmDir)\s+"
    r"|\bOpen\s+.+\s+For\s+(?:Output|Append|Binary)\b"
    r"|\b(?:Print|Put|Write)\s*#"
    r"|\.(?:SaveToFile|DeleteFile|WriteText|LoadFromFile|ReadText)\b"
    r"|\b(?:WriteUtf8Text|ReadUtf8Text|RunHidden)\s*(?:\(|\b)"
    r"|\bThisWorkbook\.Close\b"
    r"|WScript\.Shell"
    r"|\.Run\s*(?:\(|\s)",
    re.IGNORECASE,
)
MAX_MATCHES_PER_COMPONENT = 500
MAX_TOTAL_MATCHES = 5000


def _matches(lines: list[str]) -> list[tuple[int, str, str]]:
    current = "(module declarations)"
    found: list[tuple[int, str, str]] = []
    for number, line in enumerate(lines, start=1):
        match = PROCEDURE.match(line)
        if match:
            current = match.group(1)
        elif WRITE_LINE.search(line):
            found.append((number, current, line.strip()))
    return found


def _procedure_names(lines: list[str]) -> list[str]:
    """Return procedure identifiers only; never return VBA source text."""
    names: list[str] = []
    for line in lines:
        match = PROCEDURE.match(line)
        if match and match.group(1) not in names:
            names.append(match.group(1))
    return names


def _procedure_profiles(lines: list[str]) -> list[dict[str, object]]:
    """Return source-redacted procedure structure and stable hashes.

    No VBA source lines, literals, cell values, or formulas are returned.
    """
    profiles: list[dict[str, object]] = []
    current_name: str | None = None
    current_start = 0
    body: list[str] = []

    def finish(end_line: int) -> None:
        if current_name is None:
            return
        categories = {
            "assignments": 0,
            "branches": 0,
            "loops": 0,
            "date_time": 0,
            "event_target": 0,
            "calls": 0,
            "other": 0,
        }
        for item in body:
            stripped = item.strip()
            if not stripped or stripped.startswith("'"):
                continue
            if re.match(r"^(?:If|ElseIf|Else|Select Case|Case|End If)\b", stripped, re.IGNORECASE):
                categories["branches"] += 1
            elif re.match(r"^(?:For|For Each|Next|Do|Loop|While|Wend)\b", stripped, re.IGNORECASE):
                categories["loops"] += 1
            elif re.search(r"\b(?:Date|Time|CDate|IsDate|Format|Hour|Minute|Second|NormalizeTimeMS|NormalizePresenceMS)\b", stripped, re.IGNORECASE):
                categories["date_time"] += 1
            elif re.search(r"\b(?:Target|Application|EnableEvents|Intersect)\b", stripped, re.IGNORECASE):
                categories["event_target"] += 1
            elif re.match(r"^(?:Set\s+|Let\s+)?[A-Za-z_][A-Za-z0-9_.]*(?:\([^)]*\))?\s*=(?!=|>)", stripped, re.IGNORECASE):
                categories["assignments"] += 1
            elif re.match(r"^(?:Call\s+)?[A-Za-z_][A-Za-z0-9_.]*\s*(?:\(|$)", stripped, re.IGNORECASE):
                categories["calls"] += 1
            else:
                categories["other"] += 1
        normalized = "\n".join(line.rstrip() for line in body).encode("utf-8")
        profiles.append({
            "name": current_name,
            "start_line": current_start,
            "end_line": end_line,
            "body_lines": len(body),
            "source_sha256": sha256(normalized).hexdigest(),
            "categories": categories,
        })

    for number, line in enumerate(lines, start=1):
        match = PROCEDURE.match(line)
        if match:
            finish(number - 1)
            current_name = match.group(1)
            current_start = number
            body = []
            continue
        if current_name is not None:
            if re.match(r"^\s*End\s+(?:Sub|Function|Property)\b", line, re.IGNORECASE):
                finish(number)
                current_name = None
                current_start = 0
                body = []
            else:
                body.append(line)
    if current_name is not None:
        finish(len(lines))
    return profiles


def _operation(line: str) -> tuple[str, str | None]:
    if re.search(
        r"\bKill\s+|\b(?:FileCopy|MkDir|RmDir)\s+|\bOpen\s+.+\s+For\s+(?:Output|Append|Binary)\b|\b(?:Print|Put|Write)\s*#|\.(?:SaveToFile|DeleteFile)\b|\bWriteUtf8Text\s*(?:\(|\b)",
        line,
        re.IGNORECASE,
    ):
        return "external_file_write_or_delete", None
    if re.search(
        r"\bReadUtf8Text\s*(?:\(|\b)|\.(?:LoadFromFile|ReadText)\b",
        line,
        re.IGNORECASE,
    ):
        return "external_file_read", None
    if re.search(
        r"\bRunHidden\s*(?:\(|\b)|WScript\.Shell|\.Run\s*(?:\(|\s)",
        line,
        re.IGNORECASE,
    ):
        return "external_process_launch", None
    if re.search(r"\bThisWorkbook\.Close\b", line, re.IGNORECASE):
        return "workbook_close", None
    sheet = re.search(r'\b(?:Worksheets|Sheets)\s*\(\s*"([^"]+)"', line, re.IGNORECASE)
    target = sheet.group(1) if sheet else None
    if re.search(r"\.\s*Unprotect\b", line, re.IGNORECASE):
        return "protection_unprotect", target
    if re.search(r"\.\s*Protect\b", line, re.IGNORECASE):
        return "protection_protect", target
    if re.search(r"\b(?:ThisWorkbook|ActiveWorkbook)\.(?:Save|SaveAs)\b", line, re.IGNORECASE):
        return "workbook_save", target
    if re.search(r"\.\s*(?:ClearContents|Clear|Delete|Insert|Add|PasteSpecial|Copy)\b", line, re.IGNORECASE):
        return "range_or_structure_mutation", target
    if re.search(r"\.\s*(?:Value2?|Formula)\b", line, re.IGNORECASE):
        if re.search(
            r"\.\s*(?:Value2?|Formula)\b(?:\s*\([^)]*\))?\s*=",
            line,
            re.IGNORECASE,
        ):
            return "cell_value_or_formula_write", target
        return "cell_value_or_formula_access", target
    if re.search(r"\.\s*(?:Visible|Hidden|ScrollArea|EnableSelection|OnAction)\b", line, re.IGNORECASE):
        return "visibility_navigation_or_action_binding", target
    if target:
        return "worksheet_reference", target
    return "format_or_other_access_or_write", target

def _safe_label(value: object) -> str:
    return str(value).encode("unicode_escape").decode("ascii")


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
            worksheet_names = {
                str(sheet.CodeName): str(sheet.Name)
                for sheet in workbook.Worksheets
            }
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
                lines = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
                matches = _matches(lines)
                component_name = str(component.Name)
                sheet_name = worksheet_names.get(component_name)
                if not matches:
                    procedure_names = _procedure_names(lines)
                    sheet_suffix = (
                        f" worksheet_name={_safe_label(sheet_name)}"
                        if sheet_name is not None
                        else ""
                    )
                    procedures = ",".join(procedure_names) if procedure_names else "(none)"
                    component_hash = sha256("\\n".join(lines).encode("utf-8")).hexdigest()
                    print(
                        f"NO_PATTERN_MATCH_COMPONENT name={_safe_label(component_name)} "
                        f"type={int(component.Type)} lines={line_count}"
                        f"{sheet_suffix} procedures={_safe_label(procedures)} "
                        f"component_sha256={component_hash}"
                    )
                    for profile in _procedure_profiles(lines):
                        category_text = ",".join(
                            f"{key}:{value}" for key, value in profile["categories"].items()
                        )
                        print(
                            f"  PROCEDURE_PROFILE name={_safe_label(profile['name'])} "
                            f"lines={profile['start_line']}-{profile['end_line']} "
                            f"body_lines={profile['body_lines']} "
                            f"categories={category_text} "
                            f"source_sha256={profile['source_sha256']}"
                        )
                    continue
                sheet_suffix = (
                    f" worksheet_name={_safe_label(sheet_name)}"
                    if sheet_name is not None
                    else ""
                )
                print(
                    f"COMPONENT name={_safe_label(component_name)} "
                    f"type={int(component.Type)} lines={line_count} "
                    f"matched_operations={len(matches)}{sheet_suffix}"
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
