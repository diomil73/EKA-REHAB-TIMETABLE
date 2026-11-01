from dataclasses import replace
from datetime import date, datetime, timezone
import inspect

import pytest

from rehab_core.day_amendment import build_day_amendment_plan
from rehab_core.day_finalization import build_finalized_day_snapshot
from rehab_core.therapist_day_history import TherapistDayHistoryRow
from rehab_excel import day_amendment_store as module


DAY = date(2026, 10, 30)
T0 = datetime(2026, 10, 30, 15, 0, tzinfo=timezone.utc)
T1 = datetime(2026, 10, 30, 15, 5, tzinfo=timezone.utc)


def _previous():
    return build_finalized_day_snapshot(
        [],
        snapshot_id="finalized:2026-10-30:v1",
        service_date=DAY,
        finalized_at=T0,
        planned_session_count=0,
    )


def _plan():
    return build_day_amendment_plan(
        [],
        previous_snapshot=_previous(),
        replacement_snapshot_id="finalized:2026-10-30:v2",
        started_event_id="amendment-started:2026-10-30:v2",
        finalized_event_id="amendment-finalized:2026-10-30:v2",
        actor_or_source="test",
        amendment_started_at=T0,
        amendment_finalized_at=T1,
        reason_code="CORRECTION",
        summary="Correct closed day",
        planned_session_count=0,
    )


def _therapist_row(snapshot_id: str = "finalized:2026-10-30:v2"):
    return TherapistDayHistoryRow(
        snapshot_id=snapshot_id,
        service_date=DAY,
        therapist_id="T1",
        capacity_timeslots=6,
        planned_occupied_timeslots=0,
        planned_sessions=0,
        final_occupied_timeslots=0,
        final_active_sessions=0,
        replacements_received=0,
        released_by_cancellation_or_no_show=0,
        treatment_suspensions_affecting_load=0,
        approved_shared_or_over_capacity_events=0,
    )


def test_amendment_store_validates_version_chain_and_therapist_rows_before_com():
    rows = module._validate_plan_for_persistence(_plan(), [_therapist_row()])
    assert len(rows) == 1

    bad = _therapist_row("wrong-snapshot")
    with pytest.raises(module.DayAmendmentStoreError, match="replacement snapshot"):
        module._validate_plan_for_persistence(_plan(), [bad])


def test_amendment_store_rejects_broken_supersedes_link():
    plan = _plan()
    broken = replace(
        plan,
        replacement_snapshot=replace(
            plan.replacement_snapshot,
            supersedes_snapshot_id="some-other-snapshot",
        ),
    )
    with pytest.raises(module.DayAmendmentStoreError, match="supersede"):
        module._validate_plan_for_persistence(broken, [_therapist_row()])


def test_amendment_store_is_one_excel_transaction_with_single_save_boundary():
    source = inspect.getsource(module.persist_day_amendment)

    assert 'DispatchEx("Excel.Application")' in source
    assert "Workbooks.Open" in source
    assert source.count("workbook.Save()") == 1
    assert "Mutation begins only after every preflight has passed" in source
    assert "SaveChanges=False" in source

    # Never compose the transaction by invoking the three independent stores,
    # because each of those saves separately and would destroy atomicity.
    assert "append_finalized_day_snapshot(" not in source
    assert "append_therapist_day_history_rows(" not in source
    assert "append_day_history_events(" not in source


def test_amendment_store_preflights_all_history_sheets_before_mutation():
    source = inspect.getsource(module.persist_day_amendment)
    mutation_marker = source.index("# Mutation begins only after every preflight has passed.")

    prefix = source[:mutation_marker]
    suffix = source[mutation_marker:]

    assert "SNAPSHOT_SHEET" in prefix
    assert "THERAPIST_SHEET" in prefix
    assert "EVENT_SHEET" in prefix
    assert "Replacement SnapshotId already exists" in prefix
    assert "Replacement therapist/day rows already exist" in prefix
    assert "amendment EventIds already exist" in prefix

    assert "FinalizedSnapshotStatus.SUPERSEDED.value" in suffix
    assert "snapshot_values_for_com(plan.replacement_snapshot)" in suffix
    assert "therapist_values_for_com(row)" in suffix
    assert "event_values_for_com(event)" in suffix
