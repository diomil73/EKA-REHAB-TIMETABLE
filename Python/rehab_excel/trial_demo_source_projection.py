"""Combine approved PATIENTS and PROGRAM_WEEKLY_SLOTS source projections.

Pure read-only integration, not a native Excel importer or approval mechanism.
It refuses outpatients in the current inpatient-only weekly source schema.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Mapping

from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_demo_source_reconciliation import reconcile_demo_source_weekdays
from rehab_excel.trial_patient_registry_projection import (
    ConfirmedClinicalFields, TrialPatientRegistryProjection,
    project_trial_patient_registry,
)
from rehab_excel.trial_weekly_projection import (
    TrialWeeklyProjection, project_trial_weekly_rows,
)


@dataclass(frozen=True)
class TrialDemoSourceProjection:
    patient_registry: TrialPatientRegistryProjection = field(repr=False)
    weekly_slots: TrialWeeklyProjection = field(repr=False)
    patients: int
    weekday_slot_count: int


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
