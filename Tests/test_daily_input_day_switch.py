from dataclasses import replace
from datetime import date, datetime, time

import pytest

from rehab_excel.daily_input_day_switch import (
    DailyInputDayDraft, DailyInputDaySwitchError, plan_open_day_switch,
)

TODAY = date(2026, 10, 26)
TOMORROW = date(2026, 10, 27)
LATER = date(2026, 10, 28)


def draft(day, *, therapist="DEMO T1", patient="DEMO P1"):
    return DailyInputDayDraft(
        day,
        ((therapist, "ΟΧΙ", time(8), time(10), "Άδεια", "Σχόλιο"), (None,) * 6),
        ((patient, "ΝΑΙ", None, "Εξέταση", "Σχόλιο"), (None,) * 5, (None,) * 5),
    )


def switch(current, saved, target, *, finals=(), expected=None):
    return plan_open_day_switch(
        current=current, saved_drafts=saved,
        expected_current_date=expected or current.service_date,
        target_date=target, finalized_dates=finals,
    )


def test_roundtrip_between_two_open_days_preserves_raw_literal_inputs():
    today = draft(TODAY)
    tomorrow = draft(TOMORROW, therapist="DEMO T2", patient="DEMO P2")
    first = switch(today, (tomorrow,), TOMORROW)
    assert first.incoming == tomorrow
    assert first.saved_before == (tomorrow,)
    assert first.saved_after == (today, tomorrow)
    prepared = replace(tomorrow, patient_rows=(tomorrow.patient_rows[0],
                       ("DEMO P3", None, time(11), None, "incomplete draft"), (None,) * 5))
    second = switch(prepared, first.saved_after, TODAY)
    assert second.incoming == today
    late = replace(today, therapist_rows=(today.therapist_rows[0],
                   ("DEMO T4", "ΝΑΙ", None, None, "Έκτακτη άδεια", None)))
    third = switch(late, second.saved_after, TOMORROW)
    assert third.incoming == prepared
    assert third.saved_after == (late, prepared)
    assert tomorrow.patient_rows[1] == (None,) * 5


def test_new_day_has_clean_grid_and_other_dates_do_not_change():
    today = draft(TODAY)
    tomorrow = draft(TOMORROW)
    plan = switch(today, (tomorrow,), LATER)
    assert plan.incoming == today.empty_for(LATER)
    assert plan.incoming.layout == today.layout
    assert plan.saved_after == (today, tomorrow)


@pytest.mark.parametrize("case", ["stale", "same", "duplicate", "layout", "closed-source", "closed-target", "timestamp"])
def test_unsafe_switch_does_not_return_a_write_plan(case):
    today = draft(TODAY)
    tomorrow = draft(TOMORROW)
    kwargs = dict(current=today, saved_drafts=(tomorrow,),
                  expected_current_date=TODAY, target_date=TOMORROW, finalized_dates=())
    if case == "stale":
        kwargs["expected_current_date"] = LATER
    elif case == "same":
        kwargs["target_date"] = TODAY
    elif case == "duplicate":
        kwargs["saved_drafts"] = (tomorrow, tomorrow)
    elif case == "layout":
        kwargs["saved_drafts"] = (replace(tomorrow, patient_rows=tomorrow.patient_rows[:1]),)
    elif case == "closed-source":
        kwargs["finalized_dates"] = (TODAY,)
    elif case == "closed-target":
        kwargs["finalized_dates"] = (TOMORROW,)
    else:
        kwargs["target_date"] = datetime(2026, 10, 27, 0, 0)
    with pytest.raises(DailyInputDaySwitchError):
        plan_open_day_switch(**kwargs)
    assert today == draft(TODAY)
    assert tomorrow == draft(TOMORROW)


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), object(), []])
def test_unsafe_cell_value_rejected(bad):
    with pytest.raises(DailyInputDaySwitchError):
        DailyInputDayDraft(TODAY, ((bad, None, None, None, None, None),), ((None,) * 5,))
