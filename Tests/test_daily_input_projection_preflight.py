"""Synthetic orchestration tests for the independent projection preflight."""
from datetime import date, time

import pytest

import rehab_excel.daily_input_projection_preflight as preflight
from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import AbsenceKind, BaseScheduleEntry, DailyAbsence, Patient, Session
from rehab_excel.daily_input_reader import DailyInputReadResult, DailyInputSourceChange
from rehab_excel.daily_input_therapist_assignment_membership import TherapistAssignmentFact
from rehab_excel.daily_input_therapist_projection_expected import (
    TherapistDailyExpectedGrid,
)
from rehab_excel.daily_input_therapist_projection_native import TherapistDailyCellExpectation
from rehab_excel.master_projection import MasterProjectionRow
from rehab_excel.reader import WorkbookSettings


SERVICE_DATE = date(2026, 11, 23)
BASELINE_DATE = date(2026, 1, 5)
HEADERS = ("THERAPIST 1",) + ("",) * 17
PATIENTS = (Patient("P1", "FICTIONAL PATIENT", room="A1", status="Active"),)
RAW_ENTRIES = (
    BaseScheduleEntry("planner:2:ΦΘ", "P1", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1"),
)
STABLE_ENTRIES = (
    BaseScheduleEntry("assignment:00000001", "P1", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1"),
)
SESSIONS = (
    Session("assignment:00000001@2026-11-23", "P1", "THERAPIST 1",
            SERVICE_DATE, time(8, 30), "ΦΘ"),
)
STATES = (
    DailySessionState(
        session_id=SESSIONS[0].session_id,
        patient_id="P1",
        session_date=SERVICE_DATE,
        status=DailySessionStatus.THERAPIST_ABSENT,
        original_therapist_id="THERAPIST 1",
        original_time=time(8, 30),
        effective_therapist_id=None,
        effective_time=None,
        reason="fictional absence",
    ),
)
SETTINGS = WorkbookSettings(
    therapist_names=("THERAPIST 1",),
    standard_timeslots=(time(8, 30),),
    reclined_timeslots=(),
    day_patterns=("Δε",),
    therapies=("ΦΘ",),
    patient_statuses=("Active",),
    yes_no_values=("ΝΑΙ", "ΟΧΙ"),
    rooms=("A1",),
)


def _daily_input(target_date=SERVICE_DATE):
    return DailyInputReadResult(
        target_date=target_date,
        absences=(DailyAbsence(
            AbsenceKind.THERAPIST, "THERAPIST 1", SERVICE_DATE,
            reason="fictional absence",
        ),),
        therapist_rows_used=1,
        patient_rows_used=0,
        source_changes=(DailyInputSourceChange(
            "change:0001", 7, "therapist_absence", "THERAPIST 1",
            "THERAPIST 1", "fictional absence", True,
        ),),
    )


def _install_readers(monkeypatch, path, *, headers=HEADERS, stable=STABLE_ENTRIES,
                     daily_input=None, mutate=False, assignments=()):
    calls = []
    digest = preflight.file_sha256(path)
    grid = TherapistDailyExpectedGrid(
        service_date=SERVICE_DATE,
        source_workbook_sha256=digest,
        provider_headers=headers,
        cells=tuple(
            TherapistDailyCellExpectation(
                f"{chr(ord('B') + col)}{row}", None,
            )
            for row in (*range(2, 9), *range(12, 19))
            for col in range(9)
        ),
        expected_assignments=assignments,
        occupied_cells=0,
    )
    master_rows = (MasterProjectionRow(
        target_row=2, planner_row=2, patient_id="P1", room="A1",
        display_name="FICTIONAL PATIENT", infectious=False, status="Active",
    ),)

    monkeypatch.setattr(preflight, "read_patient_registry", lambda _p: PATIENTS)
    monkeypatch.setattr(preflight, "read_settings", lambda _p: SETTINGS)
    monkeypatch.setattr(preflight, "read_unified_base_schedule",
                        lambda _p, patients: RAW_ENTRIES)
    monkeypatch.setattr(preflight, "bind_stable_assignment_ids",
                        lambda _p, entries, require_complete: stable)
    monkeypatch.setattr(
        preflight, "read_recurring_change_events",
        lambda _p, entries, baseline_date: calls.append(
            ("recurring", entries, baseline_date)
        ) or (),
    )
    monkeypatch.setattr(
        preflight, "recurring_entries_on_date",
        lambda entries, baseline_date, target_date, changes: calls.append(
            ("effective", entries, baseline_date, target_date, changes)
        ) or stable,
    )
    monkeypatch.setattr(preflight, "read_patient_exits", lambda _p, patients: ())
    monkeypatch.setattr(preflight, "active_entries_on",
                        lambda entries, target_date, exits: entries)
    monkeypatch.setattr(preflight, "validate_effective_weekly_patient_slots", lambda entries: None)
    monkeypatch.setattr(preflight, "materialize_sessions_for_date",
                        lambda entries, target_date: SESSIONS)
    monkeypatch.setattr(preflight, "read_daily_input",
                        lambda _p, patients, sessions, therapist_names, allow_empty:
                        _daily_input() if daily_input is None else daily_input)
    monkeypatch.setattr(preflight, "read_daily_replacement_facts",
                        lambda _p, sessions, target_date: ())
    monkeypatch.setattr(preflight, "build_daily_session_states",
                        lambda sessions, absences, replacements, cancellations, target_date: STATES)
    monkeypatch.setattr(preflight, "read_students", lambda _p: ())
    monkeypatch.setattr(preflight, "_provider_headers", lambda _p: headers)

    def build_grid(*args, **kwargs):
        calls.append(("grid", kwargs))
        return grid

    def build_master(*args, **kwargs):
        calls.append(("master", kwargs))
        if mutate:
            path.write_bytes(path.read_bytes() + b"changed")
        return master_rows

    monkeypatch.setattr(preflight, "build_therapist_daily_expected_grid", build_grid)
    monkeypatch.setattr(preflight, "build_master_projection", build_master)
    return grid, master_rows, calls


def _source(tmp_path):
    path = tmp_path / "projection-preflight.xlsm"
    path.write_bytes(b"fictional disposable workbook bytes")
    return path


def test_preflight_assembles_stable_schedule_daily_state_and_expected_views(tmp_path, monkeypatch):
    path = _source(tmp_path)
    grid, master_rows, calls = _install_readers(monkeypatch, path)

    result = preflight.build_daily_input_projection_preflight(
        path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
    )

    assert result.source_workbook_sha256 == preflight.file_sha256(path)
    assert result.service_date == SERVICE_DATE
    assert result.baseline_date == BASELINE_DATE
    assert result.patient_count == 1
    assert result.base_assignment_count == 1
    assert result.session_count == 1
    assert result.daily_change_count == 1
    assert result.replacement_count == 0
    assert result.expected_projection == master_rows
    assert result.expected_therapist_grid == grid
    assert result.validated_changes == result.daily_input.source_changes
    assert calls[0] == ("recurring", STABLE_ENTRIES, BASELINE_DATE)
    assert calls[1][0] == "effective"
    assert calls[-2][0] == "grid"
    assert calls[-1][0] == "master"


def test_preflight_rejects_service_date_before_programme_baseline(tmp_path, monkeypatch):
    path = _source(tmp_path)
    _install_readers(monkeypatch, path)
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=BASELINE_DATE, baseline_date=SERVICE_DATE,
        )


def test_preflight_rejects_provider_header_outside_roster(tmp_path, monkeypatch):
    path = _source(tmp_path)
    headers = ("UNAPPROVED PROVIDER",) + ("",) * 17
    _install_readers(monkeypatch, path, headers=headers)
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        )


def test_preflight_rejects_unstable_assignment_identity(tmp_path, monkeypatch):
    path = _source(tmp_path)
    unstable = (
        BaseScheduleEntry("planner:2:ΦΘ", "P1", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1"),
    )
    _install_readers(monkeypatch, path, stable=unstable)
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        )


def test_preflight_rejects_daily_input_date_mismatch(tmp_path, monkeypatch):
    path = _source(tmp_path)
    _install_readers(monkeypatch, path, daily_input=_daily_input(date(2026, 11, 24)))
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        )


def test_preflight_rejects_source_mutation_during_expected_view_build(tmp_path, monkeypatch):
    path = _source(tmp_path)
    _install_readers(monkeypatch, path, mutate=True)
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        )


def test_preflight_requires_exact_independent_approval_for_shared_clock(tmp_path, monkeypatch):
    path = _source(tmp_path)
    assignments = (
        TherapistAssignmentFact(
            "assignment:00000001", "P1", "ΦΘ", "THERAPIST 1",
            SERVICE_DATE, time(8, 30),
        ),
        TherapistAssignmentFact(
            "assignment:00000002", "P2", "ΦΘ", "THERAPIST 1",
            SERVICE_DATE, time(8, 30),
        ),
    )
    _install_readers(monkeypatch, path, assignments=assignments)
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        )

    approval = (("THERAPIST 1", time(8, 30), frozenset({"P1", "P2"})),)
    grid, _master, _calls = _install_readers(
        monkeypatch, path, assignments=assignments,
    )
    result = preflight.build_daily_input_projection_preflight(
        path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
        approved_shared_memberships=approval,
    )
    assert result.approved_shared_memberships == approval
    assert result.approved_shared_clock_count == 1


def test_preflight_rejects_shared_clock_approval_with_wrong_members(tmp_path, monkeypatch):
    path = _source(tmp_path)
    assignments = (
        TherapistAssignmentFact(
            "assignment:00000001", "P1", "ΦΘ", "THERAPIST 1",
            SERVICE_DATE, time(8, 30),
        ),
        TherapistAssignmentFact(
            "assignment:00000002", "P2", "ΦΘ", "THERAPIST 1",
            SERVICE_DATE, time(8, 30),
        ),
    )
    _install_readers(monkeypatch, path, assignments=assignments)
    wrong = (("THERAPIST 1", time(8, 30), frozenset({"P1", "P3"})),)
    with pytest.raises(preflight.DailyInputProjectionPreflightError):
        preflight.build_daily_input_projection_preflight(
            path, service_date=SERVICE_DATE, baseline_date=BASELINE_DATE,
            approved_shared_memberships=wrong,
        )
