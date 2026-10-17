from datetime import time

import pytest

from rehab_core.models import BaseScheduleEntry, Patient, PatientType
from rehab_excel.outpatient_slot_validation import (
    OutpatientSlotConflictError,
    validate_outpatient_slot_compatibility,
)


def _patients():
    return [
        Patient("P-OUT", "Εξωτερικός", patient_type=PatientType.OUTPATIENT),
        Patient("P-IN", "Εσωτερικός"),
    ]


def _inpatient_entry(days: str) -> BaseScheduleEntry:
    return BaseScheduleEntry(
        base_entry_id="planner:2:ΦΘ",
        patient_id="P-IN",
        treatment="ΦΘ",
        start_time=time(10, 0),
        day_pattern=days,
        therapist_id="Γαβράς",
    )


def test_rejects_same_cell_overlapping_weekdays_between_outpatient_and_inpatient():
    with pytest.raises(OutpatientSlotConflictError, match="Use complementary weekdays"):
        validate_outpatient_slot_compatibility(
            outpatient_patient_id="P-OUT",
            therapist_id="Γαβράς",
            start_time=time(10, 0),
            day_pattern="Δε-Τε-Πα",
            existing_entries=[_inpatient_entry("Τε-Πε")],
            patients=_patients(),
        )


def test_allows_same_cell_when_weekdays_are_complementary():
    validate_outpatient_slot_compatibility(
        outpatient_patient_id="P-OUT",
        therapist_id="Γαβράς",
        start_time=time(10, 0),
        day_pattern="Δε-Πα",
        existing_entries=[_inpatient_entry("Τρ-Τε-Πε")],
        patients=_patients(),
    )


def test_allows_different_time_even_when_days_overlap():
    validate_outpatient_slot_compatibility(
        outpatient_patient_id="P-OUT",
        therapist_id="Γαβράς",
        start_time=time(10, 45),
        day_pattern="Δε-Τε-Πα",
        existing_entries=[_inpatient_entry("Δε-Τε-Πα")],
        patients=_patients(),
    )
