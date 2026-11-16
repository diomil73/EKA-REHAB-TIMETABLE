"""Strict, private outpatient projection for the existing DIOMIL schemas.

The separate institution-approved PATIENTS/OUTPATIENT_SCHEDULE XLSX is read
without mutation. No patients or values are printed, persisted in Git or
guessed from the inpatient manual A:V programme. The resulting data is an
in-memory candidate; it is NOT clinical admission or native Excel acceptance.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import time
from hashlib import sha256
from pathlib import Path

from openpyxl import load_workbook

from rehab_core.day_patterns import parse_day_pattern
from rehab_core.models import PatientType
from rehab_excel.outpatient_schedule_source import (
    OUTPATIENT_REQUIRED_HEADERS, read_outpatient_schedule,
)
from rehab_excel.patient_registry_source import read_patient_registry
from rehab_excel.trial_patient_registry_projection import PATIENTS_HEADERS
from rehab_excel.trial_private_outpatient_audit import (
    PATIENT_HEADERS, inspect_private_outpatient_source,
)

_DAY_CODES = ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")
_RECOGNIZED_PATIENT_COLUMNS = frozenset(PATIENTS_HEADERS)


class TrialPrivateOutpatientProjectionError(ValueError):
    """Only generic failure classes: never expose a private clinical value."""


@dataclass(frozen=True)
class TrialPrivateOutpatientProjection:
    source_sha256: str
    patient_rows: tuple[tuple[object, ...], ...] = field(repr=False)
    appointment_rows: tuple[tuple[str, str, str, time, str, str | None], ...] = field(repr=False)
    weekday_view_rows: tuple[tuple[object, ...], ...] = field(repr=False)
    patient_count: int = 0
    appointment_count: int = 0
    weekday_slot_count: int = 0


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _filled(value: object) -> bool:
    return value is not None and bool(str(value).strip())


def project_private_outpatient_source(
    source: str | Path,
) -> TrialPrivateOutpatientProjection:
    """Admit no partial data and bind a single independently audited snapshot.

    All fields are obtained using the existing application readers. Unknown
    populated source metadata cannot be silently discarded. The original file
    is checked again after every read, and no data is written.
    """
    path = Path(source).resolve()
    try:
        before = _digest(path)
        audit = inspect_private_outpatient_source(path)
        if not audit.source_unchanged or audit.admission_authorized:
            raise TrialPrivateOutpatientProjectionError("Outpatient audit contract invalid")
        patients = read_patient_registry(path)
        schedules = read_outpatient_schedule(path, patients=patients)
        workbook = load_workbook(path, read_only=True, data_only=False)
        try:
            sheet = workbook["PATIENTS"]
            schedule = workbook["OUTPATIENT_SCHEDULE"]
            pheaders = tuple(
                sheet.cell(1, col).value for col in range(1, sheet.max_column + 1)
            )
            sheaders = tuple(
                schedule.cell(1, col).value for col in range(1, schedule.max_column + 1)
            )
            if (
                pheaders[:6] != PATIENT_HEADERS
                or sheaders[:6] != OUTPATIENT_REQUIRED_HEADERS
                or len(set(pheaders)) != len(pheaders)
                or not _RECOGNIZED_PATIENT_COLUMNS.issuperset(
                    name for name in pheaders if name is not None
                )
            ):
                raise TrialPrivateOutpatientProjectionError(
                    "Outpatient source columns are not a supported schema"
                )
            # Additional note columns are not part of the established
            # operational schedule. Refuse populated columns rather than lose
            # clinical notes during a bulk import.
            for values in sheet.iter_rows(min_row=2, values_only=True):
                if any(
                    _filled(value) and (
                        index >= len(pheaders)
                        or pheaders[index] not in _RECOGNIZED_PATIENT_COLUMNS
                    )
                    for index, value in enumerate(values)
                ):
                    raise TrialPrivateOutpatientProjectionError(
                        "Unapproved outpatient patient metadata would be dropped"
                    )
            for values in schedule.iter_rows(min_row=2, values_only=True):
                if any(_filled(value) for value in values[6:]):
                    raise TrialPrivateOutpatientProjectionError(
                        "Unapproved outpatient appointment metadata would be dropped"
                    )
            # Keep approved SourceContext, which the standard patient reader
            # intentionally does not expose as a business model field.
            context_column = (
                pheaders.index("SourceContext") if "SourceContext" in pheaders else None
            )
            contexts = {}
            raw_patient_rows = tuple(sheet.iter_rows(min_row=2, values_only=True))
            for values in raw_patient_rows:
                if not any(_filled(value) for value in values):
                    continue
                raw_id = values[0]
                pid = (
                    str(int(raw_id)) if isinstance(raw_id, float) and raw_id.is_integer()
                    else str(raw_id).strip()
                )
                contexts[pid] = (
                    values[context_column]
                    if context_column is not None and context_column < len(values)
                    else None
                )
        finally:
            workbook.close()

        if (
            not patients or len(patients) != audit.outpatient_patients
            or len(schedules) != audit.recurring_treatments
            or any(p.patient_type != PatientType.OUTPATIENT for p in patients)
        ):
            raise TrialPrivateOutpatientProjectionError(
                "Outpatient projection differs from independent audit"
            )
        rows = tuple((
            p.patient_id, p.room, p.display_name,
            "Ν" if p.infectious else "Ο", p.status, "Εξωτερικός",
            p.hospital_mrn, p.responsible_doctor, contexts[p.patient_id],
        ) for p in patients)
        names = {p.patient_id: p.display_name for p in patients}
        appointments = []
        weekdays = []
        for entry in schedules:
            # Entry IDs originate from the physical row in the established
            # OUTPATIENT_SCHEDULE, never from a guessed display name.
            source_row = int(entry.base_entry_id.split(":")[-1])
            appointments.append((
                entry.patient_id, names[entry.patient_id], entry.treatment,
                entry.start_time, entry.day_pattern, entry.therapist_id,
            ))
            for day in sorted(parse_day_pattern(entry.day_pattern)):
                weekdays.append((
                    entry.patient_id, names[entry.patient_id], entry.treatment,
                    _DAY_CODES[int(day)], entry.start_time, entry.therapist_id,
                    f"OUTPATIENT_SCHEDULE:{source_row}",
                ))
        if len(weekdays) != audit.weekday_placements:
            raise TrialPrivateOutpatientProjectionError(
                "Outpatient daily placements differ from audited source"
            )
        if _digest(path) != before:
            raise TrialPrivateOutpatientProjectionError(
                "Private outpatient source changed during projection"
            )
        return TrialPrivateOutpatientProjection(
            source_sha256=before, patient_rows=rows,
            appointment_rows=tuple(appointments),
            weekday_view_rows=tuple(weekdays),
            patient_count=len(rows), appointment_count=len(appointments),
            weekday_slot_count=len(weekdays),
        )
    except TrialPrivateOutpatientProjectionError:
        raise
    except Exception:
        raise TrialPrivateOutpatientProjectionError(
            "Private outpatient source cannot be projected without assumptions"
        ) from None
