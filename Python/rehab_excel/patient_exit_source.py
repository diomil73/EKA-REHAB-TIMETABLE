"""Read-only, date-effective patient exits in the integrated Diomil workbook.

PATIENT_EXITS is an optional source of future schedule exclusions. It never
deletes PATIENTS or recurring entries, rewrites finalized history, or treats
discharge as a daily absence. Missing sheet means no exit events for legacy
workbooks. Any malformed *present* sheet fails closed.
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from rehab_core.lifecycle import (
    LifecyclePreview, PatientExit, PatientExitKind, active_entries_on,
    prepare_patient_exit,
)
from rehab_core.models import BaseScheduleEntry, Patient, PatientType
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule
from rehab_excel.patient_registry_source import read_patient_registry


SHEET_NAME = "PATIENT_EXITS"
HEADERS = ("PatientID", "ExitKind", "LastActiveDate", "Reason")


class PatientExitSourceError(ValueError):
    """Invalid effective patient exit source or inconsistent identity."""


def _str(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _date(value: object, *, epoch) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            converted = from_excel(value, epoch=epoch)
            if isinstance(converted, datetime):
                return converted.date()
            if isinstance(converted, date):
                return converted
        except (TypeError, ValueError, OverflowError):
            pass
        raise PatientExitSourceError("Invalid exit effective date")
    if isinstance(value, str):
        for format_ in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(value.strip(), format_).date()
            except ValueError:
                pass
    raise PatientExitSourceError("Missing or invalid exit effective date")


def read_patient_exits(
    path: str | Path, *, patients: tuple[Patient, ...] | list[Patient] | None = None,
) -> tuple[PatientExit, ...]:
    """Read and validate all explicit future stops against the live PATIENTS ID."""
    source = Path(path)
    patient_list = tuple(patients) if patients is not None else tuple(read_patient_registry(source))
    patient_by_id = {patient.patient_id: patient for patient in patient_list}
    wb = load_workbook(
        source, read_only=True, data_only=True,
        keep_vba=source.suffix.casefold() == ".xlsm",
    )
    try:
        if SHEET_NAME not in wb.sheetnames:
            return ()
        ws = wb[SHEET_NAME]
        if tuple(ws.cell(1, col).value for col in range(1, 5)) != HEADERS:
            raise PatientExitSourceError("PATIENT_EXITS schema mismatch")
        results: list[PatientExit] = []
        seen: set[str] = set()
        for rownum, values in enumerate(
            ws.iter_rows(min_row=2, max_col=4, values_only=True), start=2
        ):
            if not any(_str(value) for value in values):
                continue
            raw_pid, raw_kind, raw_last, raw_reason = values
            pid = _str(raw_pid)
            if (
                not pid or pid not in patient_by_id
                or pid.casefold() in seen
            ):
                raise PatientExitSourceError(
                    f"PATIENT_EXITS row {rownum}: unknown or duplicate PatientID"
                )
            seen.add(pid.casefold())
            try:
                kind = PatientExitKind(_str(raw_kind))
            except ValueError:
                raise PatientExitSourceError(
                    f"PATIENT_EXITS row {rownum}: unsupported exit kind"
                ) from None
            patient = patient_by_id[pid]
            expected = (
                PatientExitKind.OUTPATIENT_PROGRAM_END
                if patient.patient_type == PatientType.OUTPATIENT
                else PatientExitKind.INPATIENT_DISCHARGE
            )
            if kind is not expected:
                raise PatientExitSourceError(
                    f"PATIENT_EXITS row {rownum}: action does not match patient type"
                )
            results.append(PatientExit(
                patient_id=pid, kind=kind,
                last_active_date=_date(raw_last, epoch=wb.epoch),
                reason=_str(raw_reason),
            ))
        return tuple(results)
    finally:
        wb.close()


def read_effective_schedule_for_date(
    path: str | Path, treatment_date: date,
    *, patients: tuple[Patient, ...] | list[Patient] | None = None,
) -> tuple[BaseScheduleEntry, ...]:
    """Materialize the *date-effective* recurring source without altering it."""
    if not isinstance(treatment_date, date) or isinstance(treatment_date, datetime):
        raise PatientExitSourceError("A valid service date is required")
    registered = tuple(patients) if patients is not None else tuple(read_patient_registry(path))
    entries = read_unified_base_schedule(path, patients=registered)
    exits = read_patient_exits(path, patients=registered)
    return tuple(active_entries_on(entries, treatment_date, exits))


def preview_patient_exit_from_workbook(
    path: str | Path, *, patient_id: str,
    kind: PatientExitKind, last_active_date: date,
) -> LifecyclePreview:
    """Preview a patient exit against every live weekly recurrence, no writes."""
    registered = tuple(read_patient_registry(path))
    patients = {person.patient_id: person for person in registered}
    if patient_id not in patients:
        raise PatientExitSourceError("Unknown patient for exit preview")
    existing = read_patient_exits(path, patients=registered)
    if any(event.patient_id == patient_id for event in existing):
        raise PatientExitSourceError("Patient already has an effective exit")
    return prepare_patient_exit(
        patient=patients[patient_id], kind=kind,
        last_active_date=last_active_date,
        entries=read_unified_base_schedule(path, patients=registered),
    )
