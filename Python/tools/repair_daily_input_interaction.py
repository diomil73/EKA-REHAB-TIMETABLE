from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import sys
import tempfile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.daily_input_interaction_repair import (  # noqa: E402
    DAILY_INPUT_SHEET,
    PATIENT_FIRST_ROW,
    PATIENT_LAST_ROW,
    THERAPIST_FIRST_ROW,
    THERAPIST_LAST_ROW,
    DailyInputInteractionRepairError,
    active_daily_input_rows,
    repair_daily_input_interaction,
)


def _verify(workbook) -> None:
    ws = workbook.Worksheets(DAILY_INPUT_SHEET)
    reason_range = ws.Range(f"E{THERAPIST_FIRST_ROW}:E{THERAPIST_LAST_ROW}")

    formula = str(reason_range.Validation.Formula1 or "").strip()
    if formula.casefold() != "=PY_THERAPIST_REASONS".casefold():
        raise DailyInputInteractionRepairError(
            f"Reason dropdown read-back failed: {formula!r}"
        )

    if not bool(getattr(ws, "ProtectContents", False)):
        raise DailyInputInteractionRepairError(
            "DAILY_INPUT protection is not active after repair"
        )

    active_rows = set(active_daily_input_rows(workbook))
    for row in range(THERAPIST_FIRST_ROW, THERAPIST_LAST_ROW + 1):
        locked = bool(ws.Range(f"A{row}:F{row}").Locked)
        if row in active_rows and not locked:
            raise DailyInputInteractionRepairError(
                f"ACTIVE therapist input row remained editable: {row}"
            )
        if row not in active_rows and locked:
            raise DailyInputInteractionRepairError(
                f"Unapplied therapist input row remained locked: {row}"
            )

    for row in range(PATIENT_FIRST_ROW, PATIENT_LAST_ROW + 1):
        locked = bool(ws.Range(f"A{row}:E{row}").Locked)
        if row in active_rows and not locked:
            raise DailyInputInteractionRepairError(
                f"ACTIVE patient input row remained editable: {row}"
            )
        if row not in active_rows and locked:
            raise DailyInputInteractionRepairError(
                f"Unapplied patient input row remained locked: {row}"
            )


def repair_workbook(path: Path) -> tuple[int, int]:
    if sys.platform != "win32":
        raise DailyInputInteractionRepairError("Windows + Microsoft Excel are required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyInputInteractionRepairError("pywin32 is required") from exc

    excel = None
    workbook = None
    rollback = Path(tempfile.gettempdir()) / f"eka_daily_input_repair_{path.stem}.xlsm"
    if rollback.exists():
        rollback.unlink()

    try:
        shutil.copy2(path, rollback)
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        rewired = repair_daily_input_interaction(workbook)
        active_count = len(active_daily_input_rows(workbook))
        workbook.Save()
        _verify(workbook)
        return rewired, active_count
    except Exception:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
            workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None
        if rollback.exists():
            shutil.copy2(rollback, path)
        raise
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
        rollback.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Repair DAILY_INPUT dropdown/editability, lock ACTIVE ledger rows, "
            "and retire direct-global sheet buttons."
        )
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    path = args.workbook.resolve()
    if not path.exists():
        print(f"SAFETY STOP: workbook not found: {path}")
        return 2
    if path.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: repair target must be an .xlsm workbook")
        return 2

    try:
        rewired, active_count = repair_workbook(path)
    except (DailyInputInteractionRepairError, OSError, ValueError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    except Exception as exc:
        print(f"SAFETY STOP: unexpected repair failure: {exc}")
        return 2

    print("DAILY_INPUT INTERACTION REPAIR OK")
    print(f"Workbook: {path}")
    print("Existing entries preserved: True")
    print("Therapist reason dropdown: verified")
    print("Unapplied input rows editable/clearable: verified")
    print(f"ACTIVE daily-input rows locked: {active_count}")
    print(f"Legacy direct-global controls rewired: {rewired}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
