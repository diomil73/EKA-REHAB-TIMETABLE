from __future__ import annotations

from dataclasses import replace
from typing import Iterable, Mapping

from rehab_core.daily_state import DailySessionState
from rehab_core.models import Patient, Student

from .outpatient_presentation import build_outpatient_daily_patches
from .writeback import CellPatch, WriteIntent, WritePlan, build_write_plan


class DailyPreviewCompositionError(RuntimeError):
    """Raised when daily presentation patches cannot be merged safely."""


def merge_presentation_patches(
    base_plan: WritePlan,
    presentation_patches: Iterable[CellPatch],
) -> WritePlan:
    """Merge formatting-only patches into an existing daily write plan.

    If the base plan already writes a cell value, the presentation attributes
    are folded into that existing patch. This prevents duplicate/conflicting
    writes while preserving the authoritative value/text mutation.
    """

    ordered = list(base_plan.patches)
    index_by_key = {patch.normalized_key(): index for index, patch in enumerate(ordered)}

    for presentation in presentation_patches:
        if presentation.intent != WriteIntent.PRESENTATION:
            raise DailyPreviewCompositionError(
                "Daily presentation composer accepts PRESENTATION patches only"
            )
        key = presentation.normalized_key()
        existing_index = index_by_key.get(key)
        if existing_index is None:
            index_by_key[key] = len(ordered)
            ordered.append(presentation)
            continue

        existing = ordered[existing_index]
        ordered[existing_index] = replace(
            existing,
            fill_role=(
                presentation.fill_role
                if presentation.fill_role is not None
                else existing.fill_role
            ),
            border_role=(
                presentation.border_role
                if presentation.border_role is not None
                else existing.border_role
            ),
            source_tag=(
                f"{existing.source_tag}+{presentation.source_tag}"
                if existing.source_tag and presentation.source_tag
                else existing.source_tag or presentation.source_tag
            ),
        )

    return build_write_plan(base_plan.workbook_path, ordered)


def compose_outpatient_daily_plan(
    base_plan: WritePlan,
    *,
    states: Iterable[DailySessionState],
    patients: Iterable[Patient],
    provider_labels: Mapping[str, str] | None = None,
    students: Iterable[Student] = (),
) -> WritePlan:
    """Add outpatient light-blue presentation to an existing daily plan."""

    presentation = build_outpatient_daily_patches(
        base_plan.workbook_path,
        states=states,
        patients=patients,
        provider_labels=provider_labels,
        students=students,
    )
    return merge_presentation_patches(base_plan, presentation)
