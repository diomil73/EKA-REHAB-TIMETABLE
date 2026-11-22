"""Synthetic-only checks for safe lossless weekly trial projection."""
from datetime import time

import pytest
from openpyxl import Workbook

from rehab_core.programme_mapping import SourceWeeklyEntry
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_weekly_projection import (
    TrialWeeklyProjectionError, project_trial_weekly_rows,
)
from rehab_excel.weekly_slots_source import (
    HEADERS, read_weekly_slots_sheet,
)


def _entry(patient="DEMO-1", specialty="ΦΘ", day="ΔΕ", clock=time(9, 30), provider="DEMO STAFF", cell="C4"):
    return SourceWeeklyEntry(patient, specialty, day, clock, provider, cell)


def _plan(*entries):
    return TrialImportPlan("a" * 64, "b" * 64, 2, len(entries), tuple(entries))


def test_exact_weekly_rows_roundtrip_to_existing_reader():
    plan = _plan(_entry(), _entry("DEMO-2", day="ΤΡ", clock=time(10, 0), cell="D5"))
    projection = project_trial_weekly_rows(
        plan, confirmed_inpatient_ids=("DEMO-1", "DEMO-2")
    )
    assert projection.headers == HEADERS
    assert projection.row_count == 2
    wb = Workbook()
    ws = wb.active
    ws.title = "PROGRAM_WEEKLY_SLOTS"
    ws.append(projection.headers)
    for row in projection.rows:
        ws.append(row)
    result = read_weekly_slots_sheet(ws, inpatient_ids={"DEMO-1", "DEMO-2"})
    assert len(result) == 2
    assert result[0].day_pattern == "Δε"
    assert result[1].day_pattern == "Τρ"
    assert result[0].start_time == time(9, 30)
    wb.close()


def test_unknown_patient_id_stops_without_dropping_record():
    with pytest.raises(TrialWeeklyProjectionError, match="Unconfirmed"):
        project_trial_weekly_rows(
            _plan(_entry()), confirmed_inpatient_ids=("DEMO-OTHER",)
        )


def test_duplicate_registry_ids_stop():
    with pytest.raises(TrialWeeklyProjectionError, match="unique"):
        project_trial_weekly_rows(
            _plan(_entry()), confirmed_inpatient_ids=("DEMO-1", "DEMO-1")
        )


def test_unknown_treatment_stops():
    with pytest.raises(TrialWeeklyProjectionError, match="Unknown treatment"):
        project_trial_weekly_rows(
            _plan(_entry(specialty="UNSUPPORTED")), confirmed_inpatient_ids=("DEMO-1",)
        )


def test_missing_mandatory_therapist_stops():
    with pytest.raises(TrialWeeklyProjectionError, match="Mandatory provider"):
        project_trial_weekly_rows(
            _plan(_entry(provider=None)), confirmed_inpatient_ids=("DEMO-1",)
        )


def test_duplicate_patient_times_stop():
    with pytest.raises(TrialWeeklyProjectionError, match="collide"):
        project_trial_weekly_rows(
            _plan(_entry(), _entry(specialty="Εργο", cell="D4")),
            confirmed_inpatient_ids=("DEMO-1",)
        )


def test_plan_count_mismatch_stops():
    p = TrialImportPlan("a" * 64, "b" * 64, 1, 2, (_entry(),))
    with pytest.raises(TrialWeeklyProjectionError, match="count"):
        project_trial_weekly_rows(p, confirmed_inpatient_ids=("DEMO-1",))


def test_same_speech_therapy_source_slot_cannot_be_imported_twice():
    """Real-source pattern reproduced using fictional records ONLY.

    Duplicate weekday/time/source-cell records are not silently collapsed
    into one visit or counted as two independent appointments.
    """
    duplicate = _entry(
        specialty="Λογο", day="ΤΕ", clock=time(11, 30),
        provider=None, cell="J4",
    )
    with pytest.raises(TrialWeeklyProjectionError, match="collide"):
        project_trial_weekly_rows(
            _plan(duplicate, duplicate),
            confirmed_inpatient_ids=("DEMO-1",),
        )



def test_casefold_approved_patient_id_duplicates_stop():
    """Two case spellings of one stable PatientID cannot be two patients."""
    with pytest.raises(TrialWeeklyProjectionError, match="unique"):
        project_trial_weekly_rows(
            _plan(_entry()), confirmed_inpatient_ids=("DEMO-1", "demo-1"),
        )


def test_nontext_unhashable_approved_patient_id_is_a_safe_stop():
    """Malformed approval data must not trigger an uncontrolled set TypeError."""
    with pytest.raises(TrialWeeklyProjectionError, match="unique"):
        project_trial_weekly_rows(
            _plan(_entry()), confirmed_inpatient_ids=("DEMO-1", ["BAD-ID"]),
        )

def test_efa_same_day_distinct_times_survive_weekly_reader():
    entries = (_entry(specialty="ΕΦΑ", provider=None, cell="P4", clock=time(8)),
               _entry(specialty="ΕΦΑ", provider=None, cell="P4", clock=time(9)))
    projection = project_trial_weekly_rows(
        _plan(*entries), confirmed_inpatient_ids=("DEMO-1",))
    wb = Workbook()
    try:
        ws = wb.active
        ws.append(HEADERS)
        for row in projection.rows:
            ws.append(row)
        result = read_weekly_slots_sheet(ws, inpatient_ids={"DEMO-1"})
        assert [x.start_time for x in result] == [time(8), time(9)]
        assert all(x.day_pattern == "Δε" for x in result)
    finally:
        wb.close()


def test_exact_confirmed_physio_robotic_pair_survives_weekly_projection():
    """Fictional R6-shaped case: only the exact confirmed pair may coexist."""
    physio = _entry(specialty="ΦΘ", cell="C4")
    robotic = _entry(specialty="Ρομποτικό", cell="R4")
    plan = _plan(physio, robotic)

    with pytest.raises(TrialWeeklyProjectionError, match="collide"):
        project_trial_weekly_rows(plan, confirmed_inpatient_ids=("DEMO-1",))

    projection = project_trial_weekly_rows(
        plan,
        confirmed_inpatient_ids=("DEMO-1",),
        confirmed_combined_slots=frozenset({frozenset({physio, robotic})}),
    )
    assert projection.row_count == 2
    assert {row[1] for row in projection.rows} == {"ΦΘ", "Ρομποτικό"}


def test_stale_combined_slot_confirmation_stops_weekly_projection():
    """An exact approval must not silently authorize a changed time or day."""
    physio = _entry(specialty="ΦΘ", cell="C4")
    robotic = _entry(specialty="Ρομποτικό", cell="R4")
    changed = _entry(specialty="Ρομποτικό", cell="R4", clock=time(10, 0))
    with pytest.raises(TrialWeeklyProjectionError, match="collide"):
        project_trial_weekly_rows(
            _plan(physio, changed),
            confirmed_inpatient_ids=("DEMO-1",),
            confirmed_combined_slots=frozenset({frozenset({physio, robotic})}),
        )


def test_approved_pair_does_not_authorize_third_same_time_therapy():
    physio = _entry(specialty="ΦΘ", cell="C4")
    robotic = _entry(specialty="Ρομποτικό", cell="R4")
    speech = _entry(specialty="Λογο", cell="J4", provider=None)
    with pytest.raises(TrialWeeklyProjectionError, match="collide"):
        project_trial_weekly_rows(
            _plan(physio, robotic, speech),
            confirmed_inpatient_ids=("DEMO-1",),
            confirmed_combined_slots=frozenset({frozenset({physio, robotic})}),
        )
