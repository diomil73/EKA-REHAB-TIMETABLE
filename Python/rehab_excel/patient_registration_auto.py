from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Iterable

from rehab_core.models import Patient
from rehab_core.patient_ids import next_sequential_patient_id
from rehab_core.registration import NewPatientRequest

from .patient_registration import (
    PatientRegistrationBackend,
    PatientRegistrationPreviewReport,
    PatientRegistrationWriteError,
    create_patient_registration_preview,
)
from .master_projection_writer import (
    MasterProjectionWriteError,
    refresh_master_projection_in_place,
)
from .patient_planner_projection import (
    PatientPlannerProjectionError,
    refresh_patient_planner_projection_in_place,
)
from .patient_registry_source import read_patient_registry


def resolve_patient_id_request(
    request: NewPatientRequest,
    *,
    existing_patients: Iterable[Patient],
) -> NewPatientRequest:
    """Fill a blank PatientID with the next immutable sequential internal ID.

    Explicit IDs remain supported for legacy/import tooling. Interactive UI
    callers should leave ``patient_id`` blank so allocation is authoritative.
    """

    if request.patient_id.strip():
        return request
    return replace(
        request,
        patient_id=next_sequential_patient_id(existing_patients),
    )


def create_auto_patient_registration_preview(
    source_path: str | Path,
    output_path: str | Path,
    request: NewPatientRequest,
    *,
    backend: PatientRegistrationBackend | None = None,
    overwrite: bool = False,
) -> PatientRegistrationPreviewReport:
    existing = read_patient_registry(source_path)
    resolved = resolve_patient_id_request(request, existing_patients=existing)
    report = create_patient_registration_preview(
        source_path,
        output_path,
        resolved,
        backend=backend,
        overwrite=overwrite,
    )
    try:
        refresh_patient_planner_projection_in_place(output_path)
        refresh_master_projection_in_place(output_path)
    except (PatientPlannerProjectionError, MasterProjectionWriteError) as exc:
        try:
            Path(output_path).unlink()
        except OSError:
            pass
        raise PatientRegistrationWriteError(
            f"Patient registration projection refresh failed: {exc}"
        ) from exc
    return report
