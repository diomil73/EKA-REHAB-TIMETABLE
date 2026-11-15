"""Synthetic-only regression tests for explicit afternoon interval safety."""
from datetime import time

import pytest

from rehab_excel.afternoon_slots_source import AfternoonSlot
from rehab_excel.afternoon_shared_slot_gate import SharedSlotApproval
from rehab_excel.afternoon_interval_gate import (
    AfternoonIntervalError, AfternoonTimedSlot,
    review_afternoon_intervals, review_afternoon_timed_edit,
)
from rehab_excel.afternoon_patient_edit_plan import prepare_afternoon_patient_edit


def slot(pid, hour=16, minute=0, therapist="SYNTH-PT", weekday="ΔΕ"):
    return AfternoonSlot(pid, "ΦΘ", weekday, time(hour, minute), therapist)


def timed(pid, start=16, end=17, provider="SYNTH-PT", weekday="ΔΕ"):
    return AfternoonTimedSlot(
        slot(pid, start, therapist=provider, weekday=weekday), time(end)
    )


def consent(*members):
    return SharedSlotApproval("SYNTH-PT", "ΔΕ", time(16), frozenset(members))


def test_empty_schedule_has_no_implied_sessions():
    report = review_afternoon_intervals(())
    assert report.ready and report.slot_count == 0


def test_adjacent_intervals_do_not_overlap():
    report = review_afternoon_intervals(
        (timed("SYNTH-P1", 16, 17), timed("SYNTH-P2", 17, 18))
    )
    assert report.ready


def test_different_start_times_overlap_even_without_equal_clock():
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_intervals((
            timed("SYNTH-P1", 16, 17),
            AfternoonTimedSlot(slot("SYNTH-P2", 16, 30), time(17, 30)),
        ))


def test_same_patient_overlapping_different_therapists_is_invalid():
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_intervals((
            timed("SYNTH-P1", 16, 17),
            AfternoonTimedSlot(slot("SYNTH-P1", 16, 30, "SYNTH-OTHER"), time(17, 30)),
        ))


def test_different_therapists_overlapping_different_patients_allowed():
    report = review_afternoon_intervals((
        timed("SYNTH-P1", 16, 17),
        timed("SYNTH-P2", 16, 17, provider="SYNTH-OTHER"),
    ))
    assert report.ready


def test_same_therapist_same_start_requires_approval():
    rows = (timed("SYNTH-P1"), timed("SYNTH-P2"))
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_intervals(rows)
    report = review_afternoon_intervals(rows, approvals=(consent("SYNTH-P1", "SYNTH-P2"),))
    assert report.ready and report.shared_slot_count == 1


def test_approved_shared_group_must_have_same_end_time():
    rows = (
        timed("SYNTH-P1", 16, 17),
        timed("SYNTH-P2", 16, 18),
    )
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_intervals(rows, approvals=(consent("SYNTH-P1", "SYNTH-P2"),))


def test_stale_group_consent_invalid():
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_intervals(
            (timed("SYNTH-P1"),),
            approvals=(consent("SYNTH-P1", "SYNTH-P2"),),
        )


@pytest.mark.parametrize("end", [time(16), time(15, 59)])
def test_zero_or_negative_duration_rejected(end):
    with pytest.raises(AfternoonIntervalError):
        AfternoonTimedSlot(slot("SYNTH-P1"), end)


def test_missing_end_time_rejected():
    with pytest.raises(AfternoonIntervalError):
        AfternoonTimedSlot(slot("SYNTH-P1"), None)


def test_same_therapist_different_weekdays_are_independent():
    report = review_afternoon_intervals((
        timed("SYNTH-P1"),
        timed("SYNTH-P2", weekday="ΤΡ"),
    ))
    assert report.ready


def test_full_edit_needs_all_end_times_including_unchanged_patients():
    first, second = slot("SYNTH-P1"), slot("SYNTH-P2", 17)
    plan = prepare_afternoon_patient_edit(
        (first,), patient_id="SYNTH-P2", replacement_rows=(second,),
        known_patient_ids=("SYNTH-P1", "SYNTH-P2"),
    )
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_timed_edit(plan, explicit_end_times={second: time(18)})
    assert review_afternoon_timed_edit(
        plan, explicit_end_times={first: time(17), second: time(18)}
    ).ready


def test_stale_end_time_entry_rejected():
    planned = slot("SYNTH-P1")
    plan = prepare_afternoon_patient_edit(
        (), patient_id="SYNTH-P1", replacement_rows=(planned,),
        known_patient_ids=("SYNTH-P1", "SYNTH-P2"),
    )
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_timed_edit(
            plan, explicit_end_times={
                planned: time(17),
                slot("SYNTH-P2"): time(17),
            },
        )


def test_invalid_entry_cannot_enter_interval_gate():
    with pytest.raises(AfternoonIntervalError):
        review_afternoon_intervals((object(),))


def test_no_am_treatment_or_master_cell_in_model():
    report = review_afternoon_intervals((timed("SYNTH-P1"),))
    assert report.ready and report.slot_count == 1


def test_errors_do_not_expose_patient_identity():
    with pytest.raises(AfternoonIntervalError) as exc:
        review_afternoon_intervals((
            timed("SYNTH-P1", 16, 17),
            AfternoonTimedSlot(slot("SYNTH-P2", 16, 30), time(17, 30)),
        ))
    assert "SYNTH-P1" not in str(exc.value)
    assert "SYNTH-P2" not in str(exc.value)
