from datetime import date, time

from rehab_core.conflicts import (
    ConflictKind,
    build_operational_occurrences,
    find_operational_conflicts,
)
from rehab_core.models import AbsenceKind, DailyAbsence, ReplacementAssignment, Session


DAY = date(2026, 9, 17)
SLOT = time(12, 15)


def test_detects_therapist_double_booking():
    sessions = [
        Session("S1", "P1", "T1", DAY, SLOT),
        Session("S2", "P2", "T1", DAY, SLOT),
    ]

    conflicts = find_operational_conflicts(sessions)

    assert len(conflicts) == 1
    assert conflicts[0].kind == ConflictKind.THERAPIST_DOUBLE_BOOKING
    assert conflicts[0].subject_id == "T1"
    assert conflicts[0].occurrence_ids == ("S1", "S2")


def test_detects_patient_double_booking():
    sessions = [
        Session("S1", "P1", "T1", DAY, SLOT),
        Session("S2", "P1", "T2", DAY, SLOT),
    ]

    conflicts = find_operational_conflicts(sessions)

    assert len(conflicts) == 1
    assert conflicts[0].kind == ConflictKind.PATIENT_DOUBLE_BOOKING
    assert conflicts[0].subject_id == "P1"


def test_patient_absence_removes_occurrence_and_conflict():
    sessions = [
        Session("S1", "P1", "T1", DAY, SLOT),
        Session("S2", "P2", "T1", DAY, SLOT),
    ]
    absences = [DailyAbsence(AbsenceKind.PATIENT, "P2", DAY)]

    assert find_operational_conflicts(sessions, absences=absences) == []


def test_replaced_base_session_no_longer_blocks_original_therapist():
    sessions = [
        Session("TARGET", "P1", "T-ORIGINAL", DAY, SLOT),
        Session("S2", "P2", "T-ORIGINAL", DAY, SLOT),
    ]
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="TARGET",
        patient_id="P1",
        original_therapist_id="T-ORIGINAL",
        replacement_therapist_id="T-NEW",
        replacement_date=DAY,
        replacement_time=SLOT,
    )

    occurrences = build_operational_occurrences(sessions, replacements=[replacement])
    original_at_slot = [
        occurrence
        for occurrence in occurrences
        if occurrence.therapist_id == "T-ORIGINAL"
        and occurrence.start_time == SLOT
    ]

    assert [occurrence.source_id for occurrence in original_at_slot] == ["S2"]
    assert find_operational_conflicts(sessions, replacements=[replacement]) == []


def test_replacement_can_create_therapist_conflict_if_external_data_is_invalid():
    sessions = [
        Session("TARGET", "P1", "T-ORIGINAL", DAY, SLOT),
        Session("BUSY", "P2", "T-NEW", DAY, SLOT),
    ]
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="TARGET",
        patient_id="P1",
        original_therapist_id="T-ORIGINAL",
        replacement_therapist_id="T-NEW",
        replacement_date=DAY,
        replacement_time=SLOT,
    )

    conflicts = find_operational_conflicts(sessions, replacements=[replacement])

    assert len(conflicts) == 1
    assert conflicts[0].kind == ConflictKind.THERAPIST_DOUBLE_BOOKING
    assert conflicts[0].occurrence_ids == ("BUSY", "R1")


def test_replacement_at_new_time_releases_patient_old_time():
    sessions = [Session("TARGET", "P1", "T-ORIGINAL", DAY, SLOT)]
    replacement = ReplacementAssignment(
        replacement_id="R1",
        target_session_id="TARGET",
        patient_id="P1",
        original_therapist_id="T-ORIGINAL",
        replacement_therapist_id="T-NEW",
        replacement_date=DAY,
        replacement_time=time(13, 0),
    )

    occurrences = build_operational_occurrences(sessions, replacements=[replacement])

    assert len(occurrences) == 1
    assert occurrences[0].patient_id == "P1"
    assert occurrences[0].start_time == time(13, 0)
    assert occurrences[0].is_replacement is True
