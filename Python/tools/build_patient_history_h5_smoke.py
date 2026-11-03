from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.build_unified_preview import UnifiedBuildError, build_unified_preview  # noqa: E402
from tools.install_day_amendment_ui import (  # noqa: E402
    DayAmendmentUiInstallError,
    install_day_amendment_ui,
)


HISTORY_SHEETS = (
    "FINALIZED_DAY_SNAPSHOTS",
    "THERAPIST_DAY_HISTORY",
    "PATIENT_SESSION_HISTORY",
    "DAY_HISTORY_EVENTS",
)


class PatientHistorySmokeBuildError(RuntimeError):
    pass


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _assert_clean_history(path: Path) -> None:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        present = [name for name in HISTORY_SHEETS if name in wb.sheetnames]
    finally:
        wb.close()
    if present:
        raise PatientHistorySmokeBuildError(
            "Fresh H5 smoke unexpectedly contains pre-existing history sheets: "
            + ", ".join(present)
        )


def build_patient_history_h5_smoke(
    *,
    source: Path,
    output: Path,
    target_date: date,
) -> Path:
    source = source.resolve()
    output = output.resolve()

    if output.exists():
        raise PatientHistorySmokeBuildError(
            f"Output already exists; fresh smoke required: {output}"
        )
    if source == output:
        raise PatientHistorySmokeBuildError("Output must differ from source")

    try:
        build_unified_preview(
            source,
            output,
            target_date=target_date,
            overwrite=False,
        )
        _assert_clean_history(output)
        install_day_amendment_ui(output)
        _assert_clean_history(output)
    except (UnifiedBuildError, DayAmendmentUiInstallError, PatientHistorySmokeBuildError):
        output.unlink(missing_ok=True)
        raise
    except Exception as exc:
        output.unlink(missing_ok=True)
        raise PatientHistorySmokeBuildError(str(exc)) from exc

    return output


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Build one clean native H5 smoke workbook with current close-day and "
            "same-session day-save VBA, but no pre-existing history sheets."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--date", type=_parse_date, required=True)
    args = parser.parse_args()

    try:
        output = build_patient_history_h5_smoke(
            source=args.source,
            output=args.output,
            target_date=args.date,
        )
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("H5 PATIENT HISTORY SMOKE BUILD OK")
    print(f"Workbook: {output}")
    print(f"DAILY_INPUT date: {args.date.isoformat()}")
    print("Pre-existing history sheets: 0")
    print("Close-day action: installed")
    print("Day-save action: installed")
    print("Ready for native v1 -> v2 patient-history proof")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
