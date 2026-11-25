"""Build a complete expected THERAPIST_DAILY grid from independent daily facts.

This module reuses the established patient-centric preview renderer. It does
not read the rendered patient cells to invent expected values or AssignmentIDs.
The native COM reader consumes the returned full 126-cell map separately.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Mapping

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import (
    BaseScheduleEntry, Patient, ReplacementProviderKind, Session, Student,
)

from .authoritative_commit import file_sha256
from .daily_input_therapist_assignment_membership import TherapistAssignmentFact
from .daily_input_therapist_projection_native import TherapistDailyCellExpectation
from .layout_contract import (
    THERAPIST_DAILY_BLOCKS,
)
from .package_probe import WorkbookPackageProbe
from .patient_centric_preview import (
    _group_entries, _provider_cell, build_patient_centric_preview_plan,
    render_base_slot,
)
from .daily_input_therapist_projection_readback import _name
from .stable_recurring_ids import is_stable_assignment_id


class TherapistDailyExpectedGridError(ValueError):
    """The independent expected therapist projection is incomplete or ambiguous."""


@dataclass(frozen=True)
class TherapistDailyExpectedGrid:
    service_date: date
    source_workbook_sha256: str
    provider_headers: tuple[str, ...]
    cells: tuple[TherapistDailyCellExpectation, ...]
    expected_assignments: tuple[TherapistAssignmentFact, ...]
    occupied_cells: int


def _grid_addresses() -> tuple[str, ...]:
    from .layout_contract import cell_address
    return tuple(
        cell_address(row, col)
        for _header, first, last in THERAPIST_DAILY_BLOCKS
        for row in range(first, last + 1)
        for col in range(2, 11)
    )


def _provider_cell_for(
    workbook_path: Path,
    provider_id: str,
    start_time,
    *,
    provider_labels: Mapping[str, str],
    students: Mapping[str, Student],
    target_date: date,
    provider_kind: ReplacementProviderKind = ReplacementProviderKind.THERAPIST,
) -> str:
    try:
        return _provider_cell(
            workbook_path,
            provider_id,
            start_time,
            provider_kind=provider_kind,
            provider_labels=provider_labels,
            students=students,
            target_date=target_date,
        )
    except Exception as exc:
        raise TherapistDailyExpectedGridError(
            "Expected therapist provider/time cannot be mapped to the fixed grid"
        ) from exc


def _validate_provider_headers(
    workbook_path: Path,
    expected: tuple[str, ...],
) -> None:
    if type(expected) is not tuple or len(expected) != 18:
        raise TherapistDailyExpectedGridError("Exactly 18 independent provider headers required")
    if any(type(value) is not str or not value.strip() for value in expected):
        raise TherapistDailyExpectedGridError("Provider header expectation is blank")
    keys = []
    for value in expected:
        try:
            keys.append(_name(value))
        except Exception as exc:
            raise TherapistDailyExpectedGridError("Provider header expectation invalid") from exc
    for start in (0, 9):
        if len(set(keys[start:start + 9])) != 9:
            raise TherapistDailyExpectedGridError("Duplicate provider in a therapist header block")

    try:
        cells = WorkbookPackageProbe(workbook_path).cell_values("THERAPIST_DAILY")
    except Exception as exc:
        raise TherapistDailyExpectedGridError("Therapist grid layout cannot be read") from exc
    actual = tuple(
        str(cells.get(f"{chr(ord('B') + col)}{row}") or "").strip()
        for row in (1, 11)
        for col in range(9)
    )
    try:
        if tuple(_name(value) for value in actual) != tuple(_name(value) for value in expected):
            raise TherapistDailyExpectedGridError("Saved provider headers differ from independent roster")
    except TherapistDailyExpectedGridError:
        raise
    except Exception as exc:
        raise TherapistDailyExpectedGridError("Saved provider header is invalid") from exc


def build_therapist_daily_expected_grid(
    workbook_path: str | Path,
    *,
    service_date: date,
    expected_provider_headers: tuple[str, ...],
    base_entries: tuple[BaseScheduleEntry, ...],
    sessions: tuple[Session, ...],
    states: tuple[DailySessionState, ...],
    patients: tuple[Patient, ...],
    provider_labels: Mapping[str, str] | None = None,
    students: tuple[Student, ...] = (),
) -> TherapistDailyExpectedGrid:
    """Produce every expected patient-grid cell from independently validated facts.

    base_entries must already include stable AssignmentIDs, unified inpatient /
    granular / outpatient source, date-effective patient exits and recurring
    changes. sessions/states must be materialized for exactly service_date after
    DAILY_INPUT and active replacement facts have been independently validated.
    """
    if type(service_date) is not date:
        raise TherapistDailyExpectedGridError("Explicit service date required")
    if any(type(value) is not tuple for value in (
        base_entries, sessions, states, patients, students,
    )):
        raise TherapistDailyExpectedGridError("All source facts must be immutable tuples")
    if not base_entries or not sessions or not states or not patients:
        raise TherapistDailyExpectedGridError("Non-empty base, session, state and patient facts required")
    if any(type(item) is not BaseScheduleEntry for item in base_entries):
        raise TherapistDailyExpectedGridError("Invalid base schedule fact")
    if any(type(item) is not Session for item in sessions):
        raise TherapistDailyExpectedGridError("Invalid materialized session fact")
    if any(type(item) is not DailySessionState for item in states):
        raise TherapistDailyExpectedGridError("Invalid daily session state")
    if any(type(item) is not Patient for item in patients):
        raise TherapistDailyExpectedGridError("Invalid patient registry fact")
    if any(type(item) is not Student for item in students):
        raise TherapistDailyExpectedGridError("Invalid student registry fact")

    source = Path(workbook_path).resolve(strict=True)
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise TherapistDailyExpectedGridError("Existing saved XLSM layout required")
    before_hash = file_sha256(source)
    _validate_provider_headers(source, expected_provider_headers)

    patient_map = {item.patient_id: item for item in patients}
    if len(patient_map) != len(patients):
        raise TherapistDailyExpectedGridError("Duplicate PatientID in independent roster")
    entry_by_id = {item.base_entry_id: item for item in base_entries}
    if len(entry_by_id) != len(base_entries):
        raise TherapistDailyExpectedGridError("Duplicate recurring AssignmentID/source identity")
    if any(not is_stable_assignment_id(key) for key in entry_by_id):
        raise TherapistDailyExpectedGridError("Every recurring entry requires immutable AssignmentID")
    session_by_id = {item.session_id: item for item in sessions}
    if len(session_by_id) != len(sessions):
        raise TherapistDailyExpectedGridError("Duplicate materialized session identity")
    state_by_id = {item.session_id: item for item in states}
    if len(state_by_id) != len(states) or set(state_by_id) != set(session_by_id):
        raise TherapistDailyExpectedGridError("Daily states must cover every materialized session exactly once")

    for session in sessions:
        if session.session_date != service_date:
            raise TherapistDailyExpectedGridError("Materialized session belongs to another date")
        state = state_by_id[session.session_id]
        if (state.session_date != service_date or state.patient_id != session.patient_id
                or state.original_therapist_id != session.therapist_id
                or state.original_time != session.start_time):
            raise TherapistDailyExpectedGridError("Daily state does not bind to its exact base session")
        base_id = session.session_id.rsplit("@", 1)[0]
        entry = entry_by_id.get(base_id)
        if entry is None or entry.patient_id != session.patient_id:
            raise TherapistDailyExpectedGridError("Session does not bind to a stable base AssignmentID")
        if entry.treatment != session.treatment:
            raise TherapistDailyExpectedGridError("Session treatment differs from its base AssignmentID")

    labels = dict(provider_labels or {})
    student_map = {item.student_id: item for item in students}
    if len(student_map) != len(students):
        raise TherapistDailyExpectedGridError("Duplicate StudentID")
    therapist_ids = {
        entry.therapist_id for entry in base_entries
        if entry.therapist_id is not None
    }
    if therapist_ids & set(student_map):
        raise TherapistDailyExpectedGridError("StudentID collides with a recurring therapist identity")
    if any(type(k) is not str or not k.strip() or type(v) is not str or not v.strip()
           for k, v in labels.items()):
        raise TherapistDailyExpectedGridError("Provider label map is invalid")

    expected_values: dict[str, str | None] = {address: None for address in _grid_addresses()}
    grouped = _group_entries(base_entries)
    visible_patient_ids: dict[str, set[str]] = {}

    for (provider, start_time), entries in grouped.items():
        cell = _provider_cell_for(
            source, provider, start_time,
            provider_labels=labels, students=student_map,
            target_date=service_date,
        )
        if cell not in expected_values:
            raise TherapistDailyExpectedGridError("Base schedule maps outside the confirmed therapist grid")
        patient_ids = {entry.patient_id for entry in entries}
        names = [_name(patient_map[pid].display_name) for pid in patient_ids if pid in patient_map]
        if len(patient_ids) != len(names):
            raise TherapistDailyExpectedGridError("Base assignment references an unknown patient")
        if len(set(names)) != len(names):
            raise TherapistDailyExpectedGridError("Duplicate display names make a shared cell ambiguous")
        render = render_base_slot(provider, start_time, entries, patient_map)
        expected_values[cell] = render.text
        visible_patient_ids.setdefault(cell, set()).update(patient_ids)

    # A changed session appears in its original source cell; a replacement also
    # appears in its effective destination cell. Track identity separately from
    # the display string so duplicate visible names cannot silently alias IDs.
    for state in states:
        source_cell = _provider_cell_for(
            source, state.original_therapist_id, state.original_time,
            provider_labels=labels, students=student_map,
            target_date=service_date,
        )
        visible_patient_ids.setdefault(source_cell, set()).add(state.patient_id)
        if state.status == DailySessionStatus.REPLACED:
            if state.effective_therapist_id is None or state.effective_time is None:
                raise TherapistDailyExpectedGridError("Replacement state lacks effective provider/time")
            kind = state.effective_provider_kind or ReplacementProviderKind.THERAPIST
            destination_cell = _provider_cell_for(
                source, state.effective_therapist_id, state.effective_time,
                provider_labels=labels, students=student_map,
                target_date=service_date, provider_kind=kind,
            )
            visible_patient_ids.setdefault(destination_cell, set()).add(state.patient_id)

    for cell, patient_ids in visible_patient_ids.items():
        names = [_name(patient_map[pid].display_name) for pid in patient_ids if pid in patient_map]
        if len(patient_ids) != len(names) or len(set(names)) != len(names):
            raise TherapistDailyExpectedGridError("Duplicate or missing patient identity in a visible therapist cell")

    try:
        preview = build_patient_centric_preview_plan(
            source,
            base_entries=base_entries,
            states=states,
            sessions=sessions,
            patients=patients,
            provider_labels=labels,
            students=students,
            rebuild_multi_member_groups=True,
        )
    except Exception as exc:
        raise TherapistDailyExpectedGridError("Established patient-centric renderer could not build expected patches") from exc

    for patch in preview.write_plan.patches:
        if patch.sheet != "THERAPIST_DAILY" or patch.cell.upper() not in expected_values:
            raise TherapistDailyExpectedGridError("Patient-centric renderer produced an out-of-grid patch")
        if type(patch.value) is not str:
            raise TherapistDailyExpectedGridError("Patient-centric renderer produced a non-text cell")
        expected_values[patch.cell.upper()] = patch.value

    assignments: list[TherapistAssignmentFact] = []
    seen_assignment_ids: set[str] = set()
    for state in states:
        if state.status == DailySessionStatus.ACTIVE:
            provider = state.original_therapist_id
            start_time = state.original_time
            provider_kind = ReplacementProviderKind.THERAPIST
        elif state.status == DailySessionStatus.REPLACED:
            provider = state.effective_therapist_id
            start_time = state.effective_time
            provider_kind = state.effective_provider_kind or ReplacementProviderKind.THERAPIST
        else:
            continue
        if provider_kind != ReplacementProviderKind.THERAPIST:
            continue
        base_id = state.session_id.rsplit("@", 1)[0]
        entry = entry_by_id[base_id]
        if base_id in seen_assignment_ids:
            raise TherapistDailyExpectedGridError("Duplicate effective recurring AssignmentID")
        seen_assignment_ids.add(base_id)
        assignments.append(TherapistAssignmentFact(
            assignment_id=base_id,
            patient_id=entry.patient_id,
            treatment=entry.treatment,
            provider=str(provider),
            service_date=service_date,
            start_time=start_time,
        ))

    cells = tuple(
        TherapistDailyCellExpectation(address, expected_values[address])
        for address in sorted(expected_values)
    )
    if file_sha256(source) != before_hash:
        raise TherapistDailyExpectedGridError("Workbook bytes changed during expected-grid construction")
    return TherapistDailyExpectedGrid(
        service_date=service_date,
        source_workbook_sha256=before_hash,
        provider_headers=expected_provider_headers,
        cells=cells,
        expected_assignments=tuple(assignments),
        occupied_cells=sum(value is not None and bool(value.strip()) for value in expected_values.values()),
    )
