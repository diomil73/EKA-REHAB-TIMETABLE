from datetime import date, time

from rehab_core.models import Patient, Session, Therapist
from rehab_core.scheduling_candidates import rank_therapist_candidates


DAY = date(2026, 10, 10)
SLOTS = (time(9, 0), time(10, 0), time(11, 0))


def test_candidates_are_ranked_by_active_session_load():
    therapists = [
        Therapist("T-BUSY", "Σαρράς"),
        Therapist("T-LIGHT", "Γαύρας"),
    ]
    sessions = [
        Session("S1", "P1", "T-BUSY", DAY, time(9, 0)),
        Session("S2", "P2", "T-BUSY", DAY, time(10, 0)),
        Session("S3", "P3", "T-LIGHT", DAY, time(9, 0)),
    ]

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
        timeslots=SLOTS,
        therapists=therapists,
        sessions=sessions,
    )

    assert [candidate.therapist_id for candidate in candidates] == [
        "T-LIGHT",
        "T-BUSY",
    ]
    assert candidates[0].workload_label == "Γαύρας 1"
    assert candidates[1].workload_label == "Σαρράς 2"


def test_candidate_slots_require_both_therapist_and_patient_to_be_free():
    therapist = Therapist("T1", "Θεραπευτής")
    sessions = [
        Session("S-T", "OTHER", "T1", DAY, time(9, 0)),
        Session("S-P", "P-NEW", "T2", DAY, time(10, 0)),
    ]

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
        timeslots=SLOTS,
        therapists=[therapist],
        sessions=sessions,
    )

    assert len(candidates) == 1
    assert candidates[0].available_slots == (time(11, 0),)


def test_therapist_at_daily_capacity_is_not_proposed():
    therapist = Therapist("T1", "Full", max_daily_timeslots=2)
    sessions = [
        Session("S1", "P1", "T1", DAY, time(9, 0)),
        Session("S2", "P2", "T1", DAY, time(10, 0)),
    ]

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
        timeslots=SLOTS,
        therapists=[therapist],
        sessions=sessions,
    )

    assert candidates == ()


def test_shared_legacy_slot_does_not_create_extra_daily_capacity():
    therapist = Therapist("T1", "Full", max_daily_timeslots=3)
    sessions = [
        Session("S1", "P1", "T1", DAY, time(9, 0)),
        Session("S2", "P2", "T1", DAY, time(9, 0)),
        Session("S3", "P3", "T1", DAY, time(10, 0)),
    ]

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
        timeslots=(time(11, 0),),
        therapists=[therapist],
        sessions=sessions,
    )

    assert candidates == ()


def test_robotic_treatment_only_proposes_robotic_capable_therapists():
    therapists = [
        Therapist("T-NORMAL", "Normal", robotic_capable=False),
        Therapist("T-ROBOT", "Robot", robotic_capable=True),
    ]

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="Ρομποτικό",
        target_date=DAY,
        timeslots=SLOTS,
        therapists=therapists,
        sessions=[],
    )

    assert [candidate.therapist_id for candidate in candidates] == ["T-ROBOT"]


def test_candidate_is_omitted_when_patient_has_no_free_slot():
    therapist = Therapist("T1", "Therapist")
    sessions = [
        Session("P9", "P-NEW", "OTHER", DAY, time(9, 0)),
        Session("P10", "P-NEW", "OTHER", DAY, time(10, 0)),
        Session("P11", "P-NEW", "OTHER", DAY, time(11, 0)),
    ]

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
        timeslots=SLOTS,
        therapists=[therapist],
        sessions=sessions,
    )

    assert candidates == ()


def test_editing_can_ignore_current_session_for_patient_conflict_check():
    therapist = Therapist("T1", "Therapist")
    current = Session("CURRENT", "P-NEW", "OLD", DAY, time(10, 0))

    candidates = rank_therapist_candidates(
        patient_id="P-NEW",
        treatment="ΦΘ",
        target_date=DAY,
        timeslots=(time(10, 0),),
        therapists=[therapist],
        sessions=[current],
        ignore_session_id="CURRENT",
    )

    assert len(candidates) == 1
    assert candidates[0].available_slots == (time(10, 0),)
