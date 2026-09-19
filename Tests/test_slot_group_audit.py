from datetime import time

from rehab_core.models import BaseScheduleEntry
from rehab_core.slot_group_audit import audit_slot_groups


def _entry(entry_id: str, patient: str, days: str, therapist: str = "T"):
    return BaseScheduleEntry(
        base_entry_id=entry_id,
        patient_id=patient,
        treatment="ΦΘ",
        start_time=time(12, 15),
        day_pattern=days,
        therapist_id=therapist,
    )


def test_audit_counts_clean_pair_without_turning_pair_into_source_data():
    report = audit_slot_groups(
        [
            _entry("a", "P1", "Δε-Τε-Πα"),
            _entry("b", "P2", "Τρ-Πε"),
            BaseScheduleEntry(
                base_entry_id="c",
                patient_id="P3",
                treatment="ΦΘ",
                start_time=time(13, 0),
                day_pattern="Καθ/να",
                therapist_id="T",
            ),
        ]
    )

    assert report.base_assignment_count == 3
    assert report.group_count == 2
    assert report.clean_pair_count == 1
    assert report.multi_member_group_count == 1
    assert report.overlapping_group_count == 0


def test_audit_surfaces_same_day_overlap_for_review():
    report = audit_slot_groups(
        [
            _entry("a", "P1", "Καθ/να"),
            _entry("b", "P2", "Δε-Τε-Πα"),
        ]
    )

    assert report.group_count == 1
    assert report.clean_pair_count == 0
    assert report.overlapping_group_count == 1
    assert report.ok
