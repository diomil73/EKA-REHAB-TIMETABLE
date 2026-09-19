from datetime import date, time
from pathlib import Path

from rehab_excel.daily_input_patients import audit_daily_input_patients
from rehab_excel.daily_input_sheet import build_daily_input_spec


def test_clean_layout_spec_still_preserves_status_values_exactly():
    statuses = ("παρών", "εκτός κλινικής", "Αναβολή ΦΘ")
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α", "Β"),
        patient_names=("Π1", "Π2"),
        patient_statuses=statuses,
        timeslots=(time(8, 30), time(9, 15)),
    )
    assert spec.patient_statuses == statuses
    assert spec.patient_names == ("Π1", "Π2")


def test_authoritative_patient_audit_on_baseline_if_available():
    root = Path(__file__).resolve().parents[1]
    workbook = root / "Rehab_Center_System_v27_1.xlsm"
    if not workbook.exists():
        return
    audit = audit_daily_input_patients(workbook)
    assert audit.ok
    assert len(audit.dropdown_names) == audit.patient_rows
    assert len(set(audit.dropdown_names)) == len(audit.dropdown_names)
