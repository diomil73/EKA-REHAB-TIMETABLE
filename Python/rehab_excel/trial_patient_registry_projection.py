"""Read-only, original-source-bound PATIENTS registry projection for the trial.

Only confirmed identities and independently supplied clinical facts produce
in-memory PATIENTS rows. No original workbook writes, user-facing patient logs,
automatic PatientID reuse, infectious-status inference, or clinical admission.
The existing source A cell may mix ward and doctor; it is preserved verbatim in
SourceContext, never guessed into Room or ResponsibleDoctor.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Mapping
from zipfile import BadZipFile

from openpyxl import load_workbook

from rehab_core.models import Patient, PatientType
from rehab_excel.trial_admission_gate import CONFIRMED_MARKERS
from rehab_excel.trial_import_plan import TrialImportPlan, TrialImportPlanError
from rehab_excel.trial_source_alignment import (
    PATIENT_HEADERS, PATIENT_SHEET, PROGRAMME_SHEET,
    audit_trial_source_alignment,
)


PATIENTS_HEADERS = (
    "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση",
    "PatientType", "HospitalMRN", "ResponsibleDoctor", "SourceContext",
)


class TrialPatientRegistryProjectionError(ValueError):
    """An admission or source record is incomplete, changed or not approved."""


@dataclass(frozen=True, repr=False)
class ConfirmedClinicalFields:
    """Facts independently confirmed by an authorized reviewer, not guessed."""

    patient_type: PatientType
    infectious: bool
    room: str | None = None
    responsible_doctor: str | None = None
    hospital_mrn: str | None = None
    status: str | None = None


@dataclass(frozen=True)
class TrialPatientRegistryProjection:
    headers: tuple[str, ...]
    rows: tuple[tuple[object, ...], ...] = field(repr=False)
    patient_count: int = 0


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _source_row_number(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    try:
        integer = int(value)
        return integer if integer > 1 and float(value) == integer else None
    except (TypeError, ValueError, OverflowError):
        return None


def project_trial_patient_registry(
    plan: TrialImportPlan, *, programme: str | Path,
    review: str | Path, intake: str | Path,
    clinical_approvals: Mapping[tuple[str, int], ConfirmedClinicalFields],
) -> TrialPatientRegistryProjection:
    """Project all and only confirmed source patients into PATIENTS rows.

    The clinical_approvals key is (exact review clinic label, source Excel row).
    It must cover every reviewed patient without surplus keys. A reviewer must
    explicitly supply classification, infectious status and any known
    room/doctor/status information. An absent room/doctor stays absent.
    Both input workbooks and eligible XLSM remain strictly read-only.
    """
    if not isinstance(plan, TrialImportPlan) or plan.programme_sha256 is None:
        raise TrialPatientRegistryProjectionError("Source-bound trial plan required")
    if not isinstance(clinical_approvals, Mapping):
        raise TrialPatientRegistryProjectionError("Explicit clinical approvals required")
    p, r, c = (Path(x).resolve() for x in (programme, review, intake))
    try:
        plan.verify_sources(review=r, intake=c, programme=p)
        if not audit_trial_source_alignment(programme=p, review=r).aligned:
            raise TrialPatientRegistryProjectionError("Source alignment not approved")
        wb_source = load_workbook(p, read_only=True, data_only=True)
        try:
            wb_review = load_workbook(r, read_only=True, data_only=True)
            try:
                src = wb_source[PROGRAMME_SHEET]
                sheet = wb_review[PATIENT_SHEET]
                if tuple(sheet.cell(1, k).value for k in range(1, 7)) != PATIENT_HEADERS:
                    raise TrialPatientRegistryProjectionError("Review patient schema mismatch")
                rows: list[tuple[object, ...]] = []
                seen_ids: set[str] = set()
                seen_keys: set[tuple[str, int]] = set()
                for review_values in sheet.iter_rows(min_row=2, max_col=6, values_only=True):
                    if not any(_text(value) for value in review_values):
                        continue
                    raw_row, raw_clinic, raw_name, raw_context, raw_pid, confirm = review_values
                    row = _source_row_number(raw_row)
                    clinic, name, context = (
                        _text(raw_clinic), _text(raw_name), _text(raw_context)
                    )
                    pid = _text(raw_pid)
                    if (
                        row is None or row > src.max_row or not clinic or not name
                        or not pid or _text(confirm).upper() not in CONFIRMED_MARKERS
                    ):
                        raise TrialPatientRegistryProjectionError("Unconfirmed patient identity")
                    key = (clinic, row)
                    if key in seen_keys or pid.casefold() in seen_ids:
                        raise TrialPatientRegistryProjectionError("Duplicate patient approval")
                    seen_keys.add(key)
                    seen_ids.add(pid.casefold())
                    if (
                        _text(src.cell(row, 2).value) != name
                        or _text(src.cell(row, 1).value) != context
                    ):
                        raise TrialPatientRegistryProjectionError("Source patient identity mismatch")
                    fields = clinical_approvals.get(key)
                    if (
                        not isinstance(fields, ConfirmedClinicalFields)
                        or not isinstance(fields.patient_type, PatientType)
                        or type(fields.infectious) is not bool
                    ):
                        raise TrialPatientRegistryProjectionError(
                            "Explicit patient classification and infectious status required"
                        )
                    if fields.patient_type == PatientType.OUTPATIENT and (
                        fields.infectious or _text(fields.room)
                    ):
                        raise TrialPatientRegistryProjectionError(
                            "Outpatient clinical fields are contradictory"
                        )
                    patient = Patient(
                        patient_id=pid, display_name=name,
                        room=_text(fields.room) or None,
                        infectious=fields.infectious,
                        status=_text(fields.status) or None,
                        patient_type=fields.patient_type,
                        hospital_mrn=_text(fields.hospital_mrn) or None,
                        responsible_doctor=_text(fields.responsible_doctor) or None,
                    )
                    rows.append((
                        patient.patient_id, patient.room, patient.display_name,
                        "Ν" if patient.infectious else "Ο", patient.status,
                        "Εσωτερικός" if patient.patient_type == PatientType.INPATIENT
                        else "Εξωτερικός",
                        patient.hospital_mrn, patient.responsible_doctor,
                        context or None,
                    ))
                if (
                    not rows or len(rows) != plan.patient_count
                    or set(clinical_approvals) != seen_keys
                    or any(e.patient_id.casefold() not in seen_ids for e in plan.entries)
                ):
                    raise TrialPatientRegistryProjectionError(
                        "Registry approval scope differs from trial plan"
                    )
            finally:
                wb_review.close()
        finally:
            wb_source.close()
        plan.verify_sources(review=r, intake=c, programme=p)
    except (TrialImportPlanError, KeyError, OSError, BadZipFile, ValueError) as exc:
        if isinstance(exc, TrialPatientRegistryProjectionError):
            raise
        raise TrialPatientRegistryProjectionError(
            "SAFETY STOP: trial registry source or admission invalid"
        ) from None
    return TrialPatientRegistryProjection(PATIENTS_HEADERS, tuple(rows), len(rows))
