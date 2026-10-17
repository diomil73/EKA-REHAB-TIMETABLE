from datetime import time

from rehab_core.models import BaseScheduleEntry, Patient
from rehab_core.recurring_conflicts import find_recurring_therapist_conflicts


def test_same_therapist_time_and_overlapping_days_is_conflict():
    entries = [
        BaseScheduleEntry("E1", "P1", "ΦΘ", time(9, 15), "Δ-Τε", "Χρήστου"),
    ]
    patients = [Patient("P1", "ΔΙΑΤΣΙΝΤΟΣ"), Patient("P2", "ΠΑΠΑΓΕΩΡΓΙΟΥ")]

    conflicts = find_recurring_therapist_conflicts(
        patient_id="P2",
        therapist_id="Χρήστου",
        start_time=time(9, 15),
        day_pattern="Δ-Π",
        existing_entries=entries,
        patients=patients,
    )

    assert len(conflicts) == 1
    conflict = conflicts[0]
    assert conflict.therapist_id == "Χρήστου"
    assert conflict.existing_patient_name == "ΔΙΑΤΣΙΝΤΟΣ"
    assert conflict.new_patient_name == "ΠΑΠΑΓΕΩΡΓΙΟΥ"
    assert conflict.start_time == time(9, 15)


def test_non_overlapping_days_are_not_conflict():
    entries = [
        BaseScheduleEntry("E1", "P1", "ΦΘ", time(9, 15), "Δ-Τε", "Χρήστου"),
    ]

    conflicts = find_recurring_therapist_conflicts(
        patient_id="P2",
        therapist_id="Χρήστου",
        start_time=time(9, 15),
        day_pattern="Τρ-Π",
        existing_entries=entries,
    )

    assert conflicts == ()


def test_edit_can_ignore_current_recurring_entry():
    entries = [
        BaseScheduleEntry("CURRENT", "P2", "ΦΘ", time(9, 15), "Δ-Τε", "Χρήστου"),
    ]

    conflicts = find_recurring_therapist_conflicts(
        patient_id="P2",
        therapist_id="Χρήστου",
        start_time=time(9, 15),
        day_pattern="Δ-Τε",
        existing_entries=entries,
        ignore_base_entry_id="CURRENT",
    )

    assert conflicts == ()
