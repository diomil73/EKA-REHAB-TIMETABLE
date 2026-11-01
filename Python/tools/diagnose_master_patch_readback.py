from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

REPO_ROOT = Path(__file__).resolve().parents[2]
PREPARE_TOOL = REPO_ROOT / "Python" / "tools" / "prepare_daily_input_in_place.py"


class MasterPatchDiagnosticError(RuntimeError):
    pass


def _unescape(raw: str) -> str:
    result: list[str] = []
    escaped = False
    for ch in raw:
        if escaped:
            result.append({"t": "\t", "r": "\r", "n": "\n", "\\": "\\"}.get(ch, ch))
            escaped = False
        elif ch == "\\":
            escaped = True
        else:
            result.append(ch)
    if escaped:
        result.append("\\")
    return "".join(result)


def _prepare_payload(workbook: Path, payload: Path, response: Path) -> None:
    completed = subprocess.run(
        [
            sys.executable,
            str(PREPARE_TOOL),
            "--input",
            str(workbook),
            "--payload",
            str(payload),
            "--response",
            str(response),
        ],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        errors="replace",
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        raise MasterPatchDiagnosticError(f"Payload preparation failed: {detail}")


def _find_patch(payload: Path, cell_address: str) -> list[str]:
    for raw_line in payload.read_text(encoding="utf-8-sig").splitlines():
        if not raw_line.strip():
            continue
        parts = raw_line.split("\t")
        if len(parts) != 13 or parts[0] != "PATCH":
            continue
        decoded = [parts[0], *(_unescape(item) for item in parts[1:])]
        if decoded[1] == "MASTER_SCHEDULE" and decoded[2].upper() == cell_address.upper():
            return decoded
    raise MasterPatchDiagnosticError(f"MASTER_SCHEDULE!{cell_address} patch not found")


def _variant(value) -> str:
    try:
        return repr(value)
    except Exception:
        return "<unprintable>"


def _character_values(cell, attr: str, text_length: int) -> list[str]:
    values: list[str] = []
    for index in range(1, text_length + 1):
        try:
            font = cell.Characters(index, 1).Font
            values.append(_variant(getattr(font, attr)))
        except Exception as exc:
            values.append(f"<error:{exc}>")
    return values


def _apply_patch_probe(cell, parts: list[str]) -> None:
    intent = parts[3].lower()
    value_text = parts[4]
    strike = parts[5]
    source_tag = parts[11].lower()

    if intent == "clear":
        cell.ClearContents()
    elif intent == "value":
        cell.Value = value_text
    elif intent != "presentation":
        raise MasterPatchDiagnosticError(f"Unsupported probe intent: {intent}")

    if strike == "1":
        cell.Font.Strikethrough = True
        text_length = len(str(cell.Value or ""))
        if text_length:
            cell.Characters(1, text_length).Font.Strikethrough = True

    if source_tag == "master_daily_non_active_overlay":
        cell.Font.Color = 255  # VBA RGB(255, 0, 0)
        cell.Font.Strikethrough = True
        text_length = len(str(cell.Value or ""))
        if text_length:
            cell.Characters(1, text_length).Font.Strikethrough = True
    elif source_tag == "master_daily_patient_status_overlay":
        cell.Font.Color = 255
        cell.Font.Strikethrough = False
        text_length = len(str(cell.Value or ""))
        if text_length:
            cell.Characters(1, text_length).Font.Strikethrough = False


def diagnose(workbook_path: Path, cell_address: str) -> int:
    if sys.platform != "win32":
        raise MasterPatchDiagnosticError("Windows + Microsoft Excel required")
    if not workbook_path.exists():
        raise MasterPatchDiagnosticError(f"Workbook not found: {workbook_path}")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise MasterPatchDiagnosticError("pywin32 is required") from exc

    with tempfile.TemporaryDirectory(prefix="eka_master_patch_diag_") as temp_dir:
        temp = Path(temp_dir)
        payload = temp / "payload.tsv"
        response = temp / "response.json"
        copy_path = temp / "probe.xlsm"
        _prepare_payload(workbook_path.resolve(), payload, response)
        parts = _find_patch(payload, cell_address)
        shutil.copy2(workbook_path.resolve(), copy_path)

        print("MASTER PATCH READBACK DIAGNOSTIC")
        print(f"Source workbook: {workbook_path.resolve()}")
        print(f"Probe copy: {copy_path}")
        print(f"Patch: {parts}")

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.ScreenUpdating = False
            workbook = excel.Workbooks.Open(str(copy_path), UpdateLinks=0, ReadOnly=False, AddToMru=False)
            cell = workbook.Worksheets("MASTER_SCHEDULE").Range(cell_address)

            print("\nBEFORE")
            print(f"value={_variant(cell.Value)}")
            print(f"text={_variant(cell.Text)}")
            print(f"merge_cells={_variant(cell.MergeCells)}")
            print(f"merge_area={cell.MergeArea.Address}")
            print(f"whole_strike={_variant(cell.Font.Strikethrough)}")
            print(f"whole_color={_variant(cell.Font.Color)}")
            before_len = len(str(cell.Value or ""))
            print(f"text_length={before_len}")
            print(f"char_strike={_character_values(cell, 'Strikethrough', before_len)}")
            print(f"char_color={_character_values(cell, 'Color', before_len)}")

            _apply_patch_probe(cell, parts)

            print("\nAFTER PROBE APPLY")
            print(f"value={_variant(cell.Value)}")
            print(f"text={_variant(cell.Text)}")
            print(f"merge_cells={_variant(cell.MergeCells)}")
            print(f"merge_area={cell.MergeArea.Address}")
            print(f"whole_strike={_variant(cell.Font.Strikethrough)}")
            print(f"whole_color={_variant(cell.Font.Color)}")
            after_len = len(str(cell.Value or ""))
            strikes = _character_values(cell, "Strikethrough", after_len)
            colors = _character_values(cell, "Color", after_len)
            print(f"text_length={after_len}")
            print(f"char_strike={strikes}")
            print(f"char_color={colors}")
            print(f"all_chars_struck={all(item == 'True' for item in strikes) if strikes else _variant(cell.Font.Strikethrough) == 'True'}")
            print(f"whole_color_is_red={_variant(cell.Font.Color) == '255'}")
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
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe one MASTER patch on a temporary workbook copy.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--cell", default="D9")
    args = parser.parse_args()
    try:
        return diagnose(args.input, args.cell)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
