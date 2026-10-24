from datetime import date, time

from rehab_core.models import BaseScheduleEntry, Session
from rehab_core.reclining_adjacency import build_reclining_adjacency_penalty


DAY = date(2026, 10, 23)


def _entry(entry_id: str, patient_id: str, treatment: str, hh: int, mm: int = 0):
    return BaseScheduleEntry(
        base_entry_id=entry_id,
        patient_id=patient_id,
        treatment=treatment,
        start_time=time(hh, mm),
        day_pattern="Καθ/να",
        therapist_id=None if treatment == "Ανακλινόμενο" else "T1",
    )


def test_physio_penalty_is_distance_to_nearest_same_day_reclining_treatment():
    target = Session(
        "S1", "P1", "ABSENT", DAY, time(9, 15), treatment="ΦΘ"
    )
    penalty = build_reclining_adjacency_penalty(
        (
            _entry("PT", "P1", "ΦΘ", 9, 15),
            _entry("REC", "P1", "Ανακλινόμενο", 10, 45),
            _entry("OTHER", "P2", "Ανακλινόμενο", 8, 30),
        )
    )

    assert penalty(target, "T2", time(10, 0)) == 45
    assert penalty(target, "T2", time(9, 15)) == 90
    assert penalty(target, "T2", time(8, 30)) == 135


def test_penalty_ignores_reclining_entry_that_does_not_apply_on_service_date():
    target = Session(
        "S1", "P1", "ABSENT", DAY, time(9, 15), treatment="ΦΘ"
    )
    penalty = build_reclining_adjacency_penalty(
        (
            BaseScheduleEntry(
                base_entry_id="REC",
                patient_id="P1",
                treatment="Ανακλινόμενο",
                start_time=time(10, 0),
                day_pattern="Δευ",
            ),
        )
    )

    assert penalty(target, "T2", time(9, 15)) == 0


def test_penalty_is_zero_for_non_physio_or_patient_without_reclining():
    physio = Session(
        "S1", "P1", "ABSENT", DAY, time(9, 15), treatment="ΦΘ"
    )
    ergo = Session(
        "S2", "P1", "ABSENT", DAY, time(9, 15), treatment="Εργο"
    )
    penalty = build_reclining_adjacency_penalty(
        (_entry("REC", "P2", "Ανακλινόμενο", 10, 0),)
    )

    assert penalty(physio, "T2", time(10, 0)) == 0
    assert penalty(ergo, "T2", time(10, 0)) == 0
