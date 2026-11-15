from __future__ import annotations

from pathlib import Path
from typing import Iterable, Mapping

from rehab_core.daily_state import DailySessionState
from rehab_core.models import BaseScheduleEntry, Patient

from .master_daily_overlay import (
    _column_letters,
    _find_status_column,
    build_master_daily_overlay_plan,
)
from .master_projection import MasterProjectionRow, build_master_projection
from .writeback import CellPatch, WriteIntent, WritePlan, build_write_plan


class FullMasterDailyProjectionError(RuntimeError):
    """Raised when MASTER daily presentation cannot be rebuilt safely."""


def build_full_master_daily_plan(
    workbook_path: str | Path,
    *,
    base_entries: Iterable[BaseScheduleEntry],
    states: Iterable[DailySessionState],
    patients: Iterable[Patient],
    provider_labels: Mapping[str, str] | None = None,
    master_projection_rows: Iterable[MasterProjectionRow] | None = None,
) -> WritePlan:
    """Reset dated MASTER presentation, then reapply surviving daily overlays."""

    path = Path(workbook_path)
    patient_list = tuple(patients)
    projection = (
        tuple(master_projection_rows)
        if master_projection_rows is not None
        else build_master_projection(path, patients=patient_list)
    )
    status_column = _find_status_column(path)

    final_by_key: dict[tuple[str, str], CellPatch] = {}
    for item in projection:
        for column in range(4, 11):
            cell = f"{_column_letters(column)}{item.target_row}"
            patch = CellPatch(
                sheet="MASTER_SCHEDULE",
                cell=cell,
                intent=WriteIntent.PRESENTATION,
                source_tag="master_daily_reset",
            )
            final_by_key[patch.normalized_key()] = patch
        if status_column is not None:
            cell = f"{_column_letters(status_column)}{item.target_row}"
            patch = CellPatch(
                sheet="MASTER_SCHEDULE",
                cell=cell,
                value=item.status or "",
                source_tag="master_daily_status_reset",
            )
            final_by_key[patch.normalized_key()] = patch

    overlay = build_master_daily_overlay_plan(
        path,
        base_entries=base_entries,
        states=states,
        patients=patient_list,
        provider_labels=provider_labels,
        master_projection_rows=projection,
    )
    for patch in overlay.write_plan.patches:
        final_by_key[patch.normalized_key()] = patch

    return build_write_plan(path, final_by_key.values())
