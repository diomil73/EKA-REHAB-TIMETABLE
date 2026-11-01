from __future__ import annotations

import argparse
from pathlib import Path
import sys
import tempfile


class NativeContractInspectionError(RuntimeError):
    pass


def _module_code(vbproject, name: str) -> str:
    module = vbproject.VBComponents(name).CodeModule
    count = int(module.CountOfLines)
    return str(module.Lines(1, count)) if count else ""


def _read_log(name: str) -> str:
    path = Path(tempfile.gettempdir()) / name
    if not path.exists():
        return f"<missing: {path}>"
    return path.read_text(encoding="utf-8", errors="replace").strip()


def inspect_workbook(path: Path) -> int:
    if sys.platform != "win32":
        raise NativeContractInspectionError("Windows + Microsoft Excel required")
    if not path.exists():
        raise NativeContractInspectionError(f"Workbook not found: {path}")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise NativeContractInspectionError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(
            str(path.resolve()), UpdateLinks=0, ReadOnly=True, AddToMru=False
        )
        vbproject = workbook.VBProject
        daily = _module_code(vbproject, "modDailyInputAction")
        ledger = _module_code(vbproject, "modDailyChangeLedger")

        checks = {
            "daily.character_apply": "Private Sub ApplyCharacterStrike" in daily,
            "daily.character_verify": "Private Function VerifyCharacterStrike" in daily,
            "daily.status_normalizer": "NormalizeMasterPatientStatusPresentation master" in daily,
            "daily.status_character_clear": ".Characters(1, Len(CStr(master.Cells(row, statusCol).Value))).Font.Strikethrough = False" in daily,
            "ledger.canonical_apply": "ApplyValidatedDailyInputPayload(patchPath)" in ledger,
            "ledger.old_apply_absent": "ApplyPayloadFileInPlace" not in ledger,
            "ledger.safe_anchor": 'wsDaily.Cells.Find(What:="*"' in ledger,
            "ledger.old_broad_clear_absent": 'startRow + 200).ClearContents' not in ledger,
            "ledger.change_id_declared": "Dim changeId As String" in ledger,
        }

        print("NATIVE CONTRACT INSPECTION")
        print(f"Workbook: {path.resolve()}")
        for key, value in checks.items():
            print(f"{key}={value}")

        print("\nDAILY INPUT ERROR LOG")
        print(_read_log("eka_daily_input_last_error.log"))
        print("\nRESTORE ERROR LOG")
        print(_read_log("eka_daily_restore_last_error.log"))

        return 0 if all(checks.values()) else 3
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
    parser = argparse.ArgumentParser(description="Read-only inspection of installed native DAILY_INPUT/restore VBA contracts.")
    parser.add_argument("--input", required=True, type=Path)
    args = parser.parse_args()
    try:
        return inspect_workbook(args.input)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
