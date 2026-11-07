from datetime import date

import pytest

from rehab_core.therapist_attendance_history import (
    AttendanceHistoryError,
    AttendanceOutcome,
    TherapistAttendanceFact,
    summarize_attendance,
)


DAY1 = date(2026, 10, 29)
DAY2 = date(2026, 10, 30)


def record(day, outcome, *, snapshot="current-1", reason=None, source="approved daily roster"):
    return TherapistAttendanceFact(
        service_date=day,
        therapist_id="T1",
        snapshot_id=snapshot,
        outcome=outcome,
        source=source,
        reason_code=reason,
    )


def test_valid_two_day_ratio_with_explicit_absence_reason():
    result = summarize_attendance(
        "T1",
        (DAY1, DAY2),
        (
            record(DAY1, AttendanceOutcome.PRESENT),
            record(DAY2, AttendanceOutcome.ABSENT, snapshot="current-2", reason="SICK_LEAVE"),
        ),
        current_final_snapshot_ids={DAY1: "current-1", DAY2: "current-2"},
    )
    assert result.working_days == 2
    assert result.present_days == 1
    assert result.absent_days == 1
    assert result.presence_ratio == 0.5


def test_missing_day_is_unknown_not_zero_absences():
    with pytest.raises(AttendanceHistoryError, match="Missing verified"):
        summarize_attendance(
            "T1", (DAY1, DAY2),
            (record(DAY1, AttendanceOutcome.PRESENT),),
            current_final_snapshot_ids={DAY1: "current-1", DAY2: "current-2"},
        )


def test_superseded_snapshot_not_accepted_for_attendance():
    with pytest.raises(AttendanceHistoryError, match="CURRENT_FINAL"):
        summarize_attendance(
            "T1", (DAY1,),
            (record(DAY1, AttendanceOutcome.ABSENT, reason="SICK_LEAVE", snapshot="old"),),
            current_final_snapshot_ids={DAY1: "current-1"},
        )


def test_duplicate_date_is_unsafe():
    with pytest.raises(AttendanceHistoryError, match="Duplicate attendance fact"):
        summarize_attendance(
            "T1", (DAY1,),
            (record(DAY1, AttendanceOutcome.PRESENT), record(DAY1, AttendanceOutcome.PRESENT)),
            current_final_snapshot_ids={DAY1: "current-1"},
        )


def test_absence_requires_reason_and_all_facts_require_provenance():
    with pytest.raises(AttendanceHistoryError, match="Absence reason"):
        record(DAY1, AttendanceOutcome.ABSENT)
    with pytest.raises(AttendanceHistoryError, match="provenance"):
        record(DAY1, AttendanceOutcome.PRESENT, source=" ")


def test_nonworking_date_is_not_miscounted_as_absence():
    with pytest.raises(AttendanceHistoryError, match="Nonworking day"):
        summarize_attendance(
            "T1", (DAY1,),
            (record(DAY1, AttendanceOutcome.NONWORKING),),
            current_final_snapshot_ids={DAY1: "current-1"},
        )


def test_empty_explicit_calendar_does_not_produce_fabricated_percentage():
    result = summarize_attendance("T1", (), (), current_final_snapshot_ids={})
    assert result.working_days == 0
    assert result.presence_ratio is None
