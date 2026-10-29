from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_core.daily_statistics import calculate_daily_treatment_statistics  # noqa: E402
from rehab_core.models import SessionCancellationKind  # noqa: E402
from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.daily_input_reader import read_daily_input  # noqa: E402
from rehab_excel.daily_replacement_facts import read_daily_replacement_facts  # noqa: E402
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Verify persisted treatment-suspension fallback after global replacement solve."
    )
    parser.add_argument("--input", type=Path, required=True)
    args = parser.parse_args()

    workbook = args.input.resolve()
    if not workbook.exists():
        print(f"SAFETY STOP: workbook not found: {workbook}")
        return 2

    try:
        contract = assert_application_contract(workbook)
        patients = read_patient_registry(workbook)
        settings = read_settings(workbook)
        base_entries = read_unified_base_schedule(workbook)

        # DAILY_INPUT carries the smoke date and persisted suspension rows.
        from openpyxl import load_workbook
        from datetime import date, datetime

        wb = load_workbook(workbook, read_only=True, data_only=False, keep_vba=True)
        try:
            raw = wb["DAILY_INPUT"]["B2"].value
        finally:
            wb.close()
        if isinstance(raw, datetime):
            target_date = raw.date()
        elif isinstance(raw, date):
            target_date = raw
        else:
            text = str(raw or "").strip()
            try:
                target_date = datetime.strptime(text, "%d/%m/%Y").date()
            except ValueError:
                target_date = date.fromisoformat(text)

        sessions = materialize_sessions_for_date(base_entries, target_date)
        daily = read_daily_input(
            workbook,
            patients=patients,
            sessions=sessions,
            therapist_names=settings.therapist_names,
        )
        suspensions = tuple(
            item
            for item in daily.cancellations
            if item.kind == SessionCancellationKind.TREATMENT_SUSPENDED
        )
        if not suspensions:
            raise RuntimeError("No persisted Αναστολή θεραπείας cancellation found")

        replacements = tuple(
            read_daily_replacement_facts(
                workbook,
                sessions=sessions,
                target_date=target_date,
            )
        )
        states = build_daily_session_states(
            sessions,
            absences=daily.absences,
            replacements=replacements,
            cancellations=daily.cancellations,
            target_date=target_date,
        )
        statistics = calculate_daily_treatment_statistics(states)
        if statistics.unresolved_sessions != 0:
            raise RuntimeError(
                f"Expected zero unresolved sessions, found {statistics.unresolved_sessions}"
            )
        if statistics.treatment_suspended_sessions < 1:
            raise RuntimeError("Suspension was not counted in daily statistics")
        if not statistics.is_finalized:
            raise RuntimeError("Daily state is not finalized after suspension fallback")
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("SUSPENSION RESOLUTION SMOKE VERIFIED")
    print(f"Date: {target_date.isoformat()}")
    print(f"Persisted treatment suspensions: {len(suspensions)}")
    for item in suspensions:
        print(f"  {item.target_session_id} | {item.reason or 'Αναστολή θεραπείας'}")
    print(f"Unresolved sessions: {statistics.unresolved_sessions}")
    print(f"Treatment suspended sessions: {statistics.treatment_suspended_sessions}")
    print(f"Application contract: OK ({len(contract.component_names)} components)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
