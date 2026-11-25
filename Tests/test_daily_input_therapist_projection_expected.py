"""Synthetic expected-grid builder tests using fictional schedule identities."""
from datetime import date, time

import pytest
from openpyxl import Workbook

from rehab_core.base_schedule import materialize_sessions_for_date
from rehab_core.daily_state import build_daily_session_states
from rehab_core.models import (
    AbsenceKind, BaseScheduleEntry, DailyAbsence, Patient,
    ReplacementAssignment, ReplacementProviderKind, Student,
)
from rehab_excel.daily_input_therapist_projection_expected import (
    TherapistDailyExpectedGridError,
    build_therapist_daily_expected_grid,
)
from rehab_excel.layout_contract import THERAPIST_DAILY_BLOCKS, THERAPIST_DAILY_TIMES


SERVICE_DATE = date(2026, 11, 23)  # Monday


def _headers():
    return tuple(f"THERAPIST {i}" for i in range(1, 10)) + tuple(
        f"THERAPIST {i}" for i in range(10, 19)
    )


def _workbook(path, headers=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "THERAPIST_DAILY"
    headers = _headers() if headers is None else headers
    index = 0
    for header_row, first, last in THERAPIST_DAILY_BLOCKS:
        for col in range(2, 11):
            ws.cell(header_row, col).value = headers[index]
            index += 1
        for row, slot in zip(range(first, last + 1), THERAPIST_DAILY_TIMES, strict=True):
            ws.cell(row, 1).value = slot
    wb.save(path)
    wb.close()
    return headers


def _base_entries():
    return (
        BaseScheduleEntry("assignment:00000001", "P1", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1"),
        BaseScheduleEntry("assignment:00000002", "P2", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1"),
        BaseScheduleEntry("assignment:00000003", "P3", "Ρομποτικό", time(9, 15), "Δε", "THERAPIST 2", True),
    )


def _patients(*, duplicate=False):
    return (
        Patient("P1", "ALPHA"),
        Patient("P2", "ALPHA" if duplicate else "BETA"),
        Patient("P3", "GAMMA"),
    )


def _grid(path, *, entries=None, patients=None, states=None, sessions=None, headers=None, students=()):
    entries = _base_entries() if entries is None else entries
    patients = _patients() if patients is None else patients
    sessions = tuple(materialize_sessions_for_date(entries, SERVICE_DATE)) if sessions is None else sessions
    states = tuple(build_daily_session_states(
        sessions,
        absences=(DailyAbsence(
            AbsenceKind.PATIENT, "P1", SERVICE_DATE, reason="fictional absence",
        ),),
        target_date=SERVICE_DATE,
    )) if states is None else states
    headers = _workbook(path) if headers is None else headers
    return build_therapist_daily_expected_grid(
        path,
        service_date=SERVICE_DATE,
        expected_provider_headers=headers,
        base_entries=tuple(entries),
        sessions=tuple(sessions),
        states=tuple(states),
        patients=tuple(patients),
        students=tuple(students),
    )


def test_expected_grid_covers_all_126_cells_and_exact_active_assignment_ids(tmp_path):
    path = tmp_path / "expected-grid.xlsm"
    _workbook(path)
    result = _grid(path)
    cells = {item.cell: item.value for item in result.cells}

    assert len(result.cells) == 126
    assert result.occupied_cells == 2
    assert cells["B2"] == "ALPHA [Δε]\nBETA [Δε]"
    assert cells["C3"] == "GAMMA"
    assert {item.assignment_id for item in result.expected_assignments} == {
        "assignment:00000002", "assignment:00000003",
    }
    assert {item.patient_id for item in result.expected_assignments} == {"P2", "P3"}


def test_expected_grid_uses_effective_therapist_replacement_destination(tmp_path):
    path = tmp_path / "replacement-grid.xlsm"
    headers = _workbook(path)
    entries = _base_entries()
    sessions = tuple(materialize_sessions_for_date(entries, SERVICE_DATE))
    replacement = ReplacementAssignment(
        replacement_id="replacement:00000001",
        target_session_id=sessions[0].session_id,
        patient_id="P1",
        original_therapist_id="THERAPIST 1",
        replacement_therapist_id="THERAPIST 2",
        replacement_date=SERVICE_DATE,
        replacement_time=time(10, 0),
        reason="fictional replacement",
        replacement_provider_kind=ReplacementProviderKind.THERAPIST,
    )
    states = tuple(build_daily_session_states(
        sessions, replacements=(replacement,), target_date=SERVICE_DATE,
    ))
    result = _grid(
        path, entries=entries, patients=_patients(), sessions=sessions,
        states=states, headers=headers,
    )
    cells = {item.cell: item.value for item in result.cells}

    assert "→ THERAPIST 2 10:00" in cells["B2"]
    assert cells["C4"] == "ALPHA [Δε 23/11]"
    assignments = {item.assignment_id: item for item in result.expected_assignments}
    assert assignments["assignment:00000001"].provider == "THERAPIST 2"
    assert assignments["assignment:00000001"].start_time == time(10, 0)


def test_expected_grid_rejects_ambiguous_duplicate_patient_display_names(tmp_path):
    path = tmp_path / "ambiguous-grid.xlsm"
    _workbook(path)
    with pytest.raises(TherapistDailyExpectedGridError):
        _grid(path, patients=_patients(duplicate=True))


def test_expected_grid_rejects_unstable_recurring_identity(tmp_path):
    path = tmp_path / "unstable-grid.xlsm"
    headers = _workbook(path)
    entries = (
        BaseScheduleEntry("planner:2:ΦΘ", "P1", "ΦΘ", time(8, 30), "Δε", "THERAPIST 1"),
    )
    patients = (Patient("P1", "ALPHA"),)
    sessions = tuple(materialize_sessions_for_date(entries, SERVICE_DATE))
    states = tuple(build_daily_session_states(
        sessions,
        absences=(DailyAbsence(AbsenceKind.PATIENT, "P1", SERVICE_DATE),),
        target_date=SERVICE_DATE,
    ))
    with pytest.raises(TherapistDailyExpectedGridError):
        _grid(path, entries=entries, patients=patients, sessions=sessions,
              states=states, headers=headers)


def test_expected_grid_rejects_provider_roster_drift(tmp_path):
    path = tmp_path / "header-drift.xlsm"
    headers = list(_workbook(path))
    headers[0] = "WRONG PROVIDER"
    with pytest.raises(TherapistDailyExpectedGridError):
        _grid(path, headers=tuple(headers))


def test_expected_grid_handles_student_replacement_with_canonical_student_id(tmp_path):
    path = tmp_path / "student-grid.xlsm"
    headers = list(_headers())
    headers[-1] = "STUDENT A"
    _workbook(path, tuple(headers))
    entries = _base_entries()
    sessions = tuple(materialize_sessions_for_date(entries, SERVICE_DATE))
    student = Student(
        "S001", "STUDENT A", date(2026, 1, 1), date(2027, 1, 1),
    )
    replacement = ReplacementAssignment(
        replacement_id="replacement:student:0001",
        target_session_id=sessions[0].session_id,
        patient_id="P1",
        original_therapist_id="THERAPIST 1",
        replacement_therapist_id="S001",
        replacement_date=SERVICE_DATE,
        replacement_time=time(10, 0),
        reason="fictional student cover",
        replacement_provider_kind=ReplacementProviderKind.STUDENT,
    )
    states = tuple(build_daily_session_states(
        sessions, replacements=(replacement,), target_date=SERVICE_DATE,
    ))
    result = _grid(
        path, entries=entries, patients=_patients(), sessions=sessions,
        states=states, headers=tuple(headers), students=(student,),
    )
    cells = {item.cell: item.value for item in result.cells}
    assert cells["J14"] == "ALPHA [Δε 23/11]"
    assert "→ STUDENT A 10:00" in cells["B2"]
    assert {item.assignment_id for item in result.expected_assignments} == {
        "assignment:00000002", "assignment:00000003",
    }
