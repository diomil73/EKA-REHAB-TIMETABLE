from __future__ import annotations

from pathlib import Path
from typing import Iterable, Mapping

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import BaseScheduleEntry, Patient, Session, Student

from .layout_contract import resolve_therapist_daily_cell
from .patient_centric_preview import (
    PatientCentricPreviewError,
    _line_runs,
    build_patient_centric_preview_plan,
    render_base_slot,
)
from .writeback import CellPatch, WriteIntent, WritePlan, build_write_plan


class FullDailyProjectionError(RuntimeError):
    """Raised when THERAPIST_DAILY cannot be rebuilt authoritatively."""


_GRID_CELLS = tuple(
    f"{column}{row}"
    for row in (*range(2, 9), *range(12, 19))
    for column in "BCDEFGHIJ"
)


def build_full_therapist_daily_plan(
    workbook_path: str | Path,
    *,
    base_entries: Iterable[BaseScheduleEntry],
    states: Iterable[DailySessionState],
    sessions: Iterable[Session],
    patients: Iterable[Patient],
    provider_labels: Mapping[str, str] | None = None,
    students: Iterable[Student] = (),
) -> WritePlan:
    """Rebuild the complete therapist grid from base truth plus dated overlays.

    The plan writes every confirmed THERAPIST_DAILY grid cell. Base recurring
    slots are reconstructed patient-centrically; empty cells are cleared so a
    removed replacement cannot leave a stale destination. Non-active/replaced
    dated states are then overlaid using the same patient-centric renderer used
    by normal DAILY_INPUT application.
    """

    path = Path(workbook_path)
    entry_list = tuple(base_entries)
    state_list = tuple(states)
    patient_list = tuple(patients)
    session_list = tuple(sessions)
    patient_map = {patient.patient_id: patient for patient in patient_list}

    grouped: dict[tuple[str, object], list[BaseScheduleEntry]] = {}
    for entry in entry_list:
        if entry.therapist_id is None:
            continue
        grouped.setdefault((entry.therapist_id, entry.start_time), []).append(entry)

    by_cell: dict[str, CellPatch] = {
        cell: CellPatch(
            sheet="THERAPIST_DAILY",
            cell=cell,
            intent=WriteIntent.CLEAR,
            source_tag="full_daily_base_clear",
        )
        for cell in _GRID_CELLS
    }

    for (therapist_id, slot_time), entries in grouped.items():
        try:
            cell = resolve_therapist_daily_cell(path, therapist_id, slot_time)
        except Exception as exc:
            raise FullDailyProjectionError(str(exc)) from exc
        render = render_base_slot(therapist_id, slot_time, entries, patient_map)
        by_cell[cell] = CellPatch(
            sheet="THERAPIST_DAILY",
            cell=cell,
            value=render.text,
            wrap_text=True,
            min_font_size=10,
            text_runs=_line_runs(render.lines),
            source_tag="full_daily_base",
        )

    changed = tuple(
        state for state in state_list if state.status != DailySessionStatus.ACTIVE
    )
    if changed:
        try:
            overlay = build_patient_centric_preview_plan(
                path,
                base_entries=entry_list,
                states=state_list,
                sessions=session_list,
                patients=patient_list,
                provider_labels=provider_labels,
                students=students,
                rebuild_multi_member_groups=True,
            )
        except PatientCentricPreviewError as exc:
            raise FullDailyProjectionError(str(exc)) from exc
        for patch in overlay.write_plan.patches:
            by_cell[patch.cell.upper()] = patch

    ordered = [by_cell[cell] for cell in _GRID_CELLS]
    extras = [
        patch
        for cell, patch in sorted(by_cell.items())
        if cell not in set(_GRID_CELLS)
    ]
    return build_write_plan(path, (*ordered, *extras))
