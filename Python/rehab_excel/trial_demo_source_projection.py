"""Combine approved PATIENTS and PROGRAM_WEEKLY_SLOTS source projections.

Pure read-only integration, not a native Excel importer or approval mechanism.
It refuses outpatients in the current inpatient-only weekly source schema.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from datetime import time
from typing import Mapping
from rehab_core.recurring_conflicts import _provider_key

from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_demo_source_reconciliation import reconcile_demo_source_weekdays
from rehab_excel.trial_patient_registry_projection import (
    ConfirmedClinicalFields, TrialPatientRegistryProjection,
    project_trial_patient_registry,
)
from rehab_excel.trial_private_outpatient_projection import TrialPrivateOutpatientProjection
from rehab_excel.trial_weekly_projection import (
    TrialWeeklyProjection, project_trial_weekly_rows,
)


@dataclass(frozen=True)
class TrialDemoSourceProjection:
    patient_registry: TrialPatientRegistryProjection = field(repr=False)
    weekly_slots: TrialWeeklyProjection = field(repr=False)
    patients: int
    weekday_slot_count: int
    outpatient: TrialPrivateOutpatientProjection | None = field(default=None, repr=False)


def merge_trial_demo_outpatients(
    inpatient: TrialDemoSourceProjection,
    outpatient: TrialPrivateOutpatientProjection,
    *,
    shared_approvals: Mapping[tuple[str, str, time], frozenset[str]] | None = None,
) -> TrialDemoSourceProjection:
    """Join independent approved sources without repeating IDs or mixing storage.

    Inpatient weekly slots remain in PROGRAM_WEEKLY_SLOTS, whereas outpatient
    rows use the existing OUTPATIENT_SCHEDULE. The result is still only an
    in-memory read-only proposal, never a clinical approval.
    """
    if not isinstance(inpatient, TrialDemoSourceProjection) or not isinstance(
        outpatient, TrialPrivateOutpatientProjection
    ) or inpatient.outpatient is not None:
        raise ValueError("Outpatient source join requires one clean original projection")
    original = inpatient.patient_registry.rows
    if any(row[5] != "Εσωτερικός" for row in original):
        raise ValueError("Inpatient source contains non-inpatients")
    if any(row[5] != "Εξωτερικός" for row in outpatient.patient_rows):
        raise ValueError("Outpatient source contains non-outpatients")
    rows = original + outpatient.patient_rows
    keys = [str(row[0]).strip().casefold() for row in rows]
    if not keys or any(not key for key in keys) or len(keys) != len(set(keys)):
        raise ValueError("Duplicate or missing patient identities across sources")
    if len(outpatient.patient_rows) != outpatient.patient_count or (
        len(outpatient.appointment_rows) != outpatient.appointment_count
        or len(outpatient.weekday_view_rows) != outpatient.weekday_slot_count
    ):
        raise ValueError("Outpatient source counts disagree")
    # Inpatient-only groups have already been checked by admission. For any
    # group that contains an outpatient, require the separately approved exact
    # membership of the merged slot. No approval is inferred from the sources.
    # Reuse the established native-recurring conflict normalization, including
    # Greek accents, final sigma, spaces and periods. Do not infer other aliases.
    def provider_slot(provider, day, clock):
        key = _provider_key(provider)
        if not key:
            raise ValueError("Provider label cannot be resolved by recurring conflict rules")
        return (key, day, clock)

    all_members = {}
    outpatient_provider_slots = set()
    for pid, _treatment, day, clock, provider, _source in inpatient.weekly_slots.rows:
        if provider:
            key = provider_slot(provider, day, clock)
            all_members.setdefault(key, set()).add(pid)
    outpatient_patient_slots = set()
    for pid, _name, _treatment, day, clock, provider, _source in outpatient.weekday_view_rows:
        patient_slot = (str(pid).strip().casefold(), day, clock)
        if patient_slot in outpatient_patient_slots:
            raise ValueError("Outpatient simultaneous treatments require independent review")
        outpatient_patient_slots.add(patient_slot)
        if not provider:
            continue
        key = provider_slot(provider, day, clock)
        all_members.setdefault(key, set()).add(pid)
        outpatient_provider_slots.add(key)

    approved = {}
    for key, members in (shared_approvals or {}).items():
        if (
            not isinstance(key, tuple) or len(key) != 3
            or not isinstance(key[0], str)
            or key[1] not in ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")
            or not isinstance(key[2], time)
            or not isinstance(members, frozenset) or len(members) < 2
            or any(not isinstance(pid, str) or not pid.strip() for pid in members)
        ):
            raise ValueError("Invalid explicit shared-slot approval")
        normalized = provider_slot(*key)
        if normalized in approved:
            raise ValueError("Duplicate explicit shared-slot approval")
        approved[normalized] = members
    shared = {
        key: frozenset(all_members[key])
        for key in outpatient_provider_slots if len(all_members[key]) > 1
    }
    if any(approved.get(key) != members for key, members in shared.items()):
        raise ValueError("Outpatient provider slot requires explicit shared-slot approval")
    if approved != shared:
        raise ValueError("Stale or mismatched explicit shared-slot approval")

    registry = TrialPatientRegistryProjection(
        inpatient.patient_registry.headers, rows, len(rows)
    )
    return TrialDemoSourceProjection(
        patient_registry=registry, weekly_slots=inpatient.weekly_slots,
        patients=len(rows),
        weekday_slot_count=inpatient.weekday_slot_count + outpatient.weekday_slot_count,
        outpatient=outpatient,
    )


def project_trial_demo_sources(
    plan: TrialImportPlan, *, programme: str | Path,
    review: str | Path, intake: str | Path,
    clinical_approvals: Mapping[tuple[str, int], ConfirmedClinicalFields],
) -> TrialDemoSourceProjection:
    """Produce mutually consistent source-bound PATIENTS and weekly rows.

    A downstream native writer must independently verify hashes again, use a
    never-existing output path, preserve VBA/history, and perform native readback.
    Neither this function nor any green synthetic test authorizes clinical use.
    """
    # Identity/source-cell alignment alone is insufficient: the independently
    # parsed source programme must equal every admitted weekly placement.
    reconcile_demo_source_weekdays(
        plan, programme=programme, review=review, intake=intake,
    )
    patient_registry = project_trial_patient_registry(
        plan, programme=programme, review=review, intake=intake,
        clinical_approvals=clinical_approvals,
    )
    inpatient_ids = tuple(
        str(row[0]) for row in patient_registry.rows if row[5] == "Εσωτερικός"
    )
    weekly = project_trial_weekly_rows(
        plan, confirmed_inpatient_ids=inpatient_ids,
    )
    if patient_registry.patient_count != plan.patient_count or weekly.row_count != plan.weekday_slot_count:
        raise ValueError("Trial patient/week projection count mismatch")
    return TrialDemoSourceProjection(
        patient_registry=patient_registry, weekly_slots=weekly,
        patients=patient_registry.patient_count,
        weekday_slot_count=weekly.row_count,
    )
