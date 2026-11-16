"""Private read-only audit of EXISTING outpatient Excel schemas.

No new outpatient forms or cell layouts. The input is a separate, authorized,
current .xlsx export with the established PATIENTS (patient registry) and
OUTPATIENT_SCHEDULE (recurring appointments) sheet headers. This audit is NOT
source approval or an importer. Formula-only cells, duplicates, cross-type
links, weekend patterns and inconsistent source counts stop the audit.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path

from openpyxl import load_workbook

from rehab_core.day_patterns import parse_day_pattern
from rehab_core.models import PatientType
from rehab_excel.outpatient_schedule_source import (
    OUTPATIENT_REQUIRED_HEADERS, _patient_id, read_outpatient_schedule,
)
from rehab_excel.patient_registry_source import read_patient_registry


PATIENT_HEADERS = (
    "PatientID", "Θάλαμος", "Ασθενής", "Λοιμώδης", "Κατάσταση", "PatientType",
)


class PrivateOutpatientSourceError(ValueError):
    """Unsafe local outpatient source; never contain private source facts."""


@dataclass(frozen=True)
class PrivateOutpatientAudit:
    outpatient_patients: int
    recurring_treatments: int
    weekday_placements: int
    distinct_treatments: int
    source_unchanged: bool
    admission_authorized: bool = False


def _hash(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _nonblank(value: object) -> bool:
    return value is not None and bool(str(value).strip())


def _check_raw_sheet(ws, expected: tuple[str, ...]) -> int:
    """Validate exactly the existing source headers, and no formula-only rows."""
    headers = tuple(ws.cell(1, col).value for col in range(1, len(expected) + 1))
    if headers != expected:
        raise PrivateOutpatientSourceError("Private outpatient sheet schema differs")
    populated = 0
    for values in ws.iter_rows(min_row=2, max_col=len(expected)):
        if not any(_nonblank(cell.value) for cell in values):
            continue
        populated += 1
        if any(cell.data_type == "f" for cell in values):
            raise PrivateOutpatientSourceError(
                "Formula-only private outpatient source has no approved cached values"
            )
    return populated


def inspect_private_outpatient_source(path: str | Path) -> PrivateOutpatientAudit:
    """Reuse existing production PATIENTS/OUTPATIENT_SCHEDULE rules, no writes."""
    source = Path(path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsx":
        raise PrivateOutpatientSourceError("Private outpatient XLSX is required")
    if any((folder / ".git").exists() for folder in (source.parent, *source.parent.parents)):
        raise PrivateOutpatientSourceError("Private outpatient sources cannot reside in Git")
    before = _hash(source)
    wb = None
    try:
        wb = load_workbook(source, read_only=True, data_only=False)
        if "PATIENTS" not in wb.sheetnames or "OUTPATIENT_SCHEDULE" not in wb.sheetnames:
            raise PrivateOutpatientSourceError("Private outpatient source sheets missing")
        patient_rows = _check_raw_sheet(wb["PATIENTS"], PATIENT_HEADERS)
        recurring_rows = _check_raw_sheet(
            wb["OUTPATIENT_SCHEDULE"], OUTPATIENT_REQUIRED_HEADERS,
        )
        if not patient_rows or not recurring_rows:
            raise PrivateOutpatientSourceError("Real outpatient source has no appointments")
        # Reject a row that has a name but no PatientID, or vice versa.
        ids = set()
        for cells in wb["PATIENTS"].iter_rows(min_row=2, max_col=6):
            if not any(_nonblank(c.value) for c in cells):
                continue
            pid, room, name, infectious, status, kind = (c.value for c in cells)
            normalized_id = _patient_id(pid)
            if (
                not normalized_id or not _nonblank(name)
                or normalized_id.casefold() in ids
                or not _nonblank(kind)
            ):
                raise PrivateOutpatientSourceError(
                    "Outpatient registry has incomplete or duplicate identities"
                )
            ids.add(normalized_id.casefold())
            # The existing patient registry reader owns all legitimate
            # inpatient/outpatient label variants; no second parser here.
            if _nonblank(room) or str(infectious or "").strip().upper() in ("Ν", "N", "YES", "TRUE", "1"):
                raise PrivateOutpatientSourceError(
                    "Outpatient clinical row is inconsistent with approved type"
                )
        if len(ids) != patient_rows:
            raise PrivateOutpatientSourceError("Outpatient registry does not match rows")
        if not source.is_file() or _hash(source) != before:
            raise PrivateOutpatientSourceError(
                "SAFETY STOP: source changed during raw outpatient inspection"
            )
    except PrivateOutpatientSourceError:
        raise
    except Exception:
        raise PrivateOutpatientSourceError(
            "Unable to inspect private outpatient source without assumptions"
        ) from None
    finally:
        if wb is not None:
            wb.close()

    try:
        patients = read_patient_registry(source)
        if (
            len(patients) != patient_rows
            or any(patient.patient_type != PatientType.OUTPATIENT for patient in patients)
        ):
            raise PrivateOutpatientSourceError(
                "Outpatient source registry differs from the production reader"
            )
        recurring = read_outpatient_schedule(source, patients=patients)
        if len(recurring) != recurring_rows:
            raise PrivateOutpatientSourceError(
                "Outpatient appointments were silently skipped by the production reader"
            )
        seen = set()
        days = 0
        counts = {patient.patient_id: 0 for patient in patients}
        for entry in recurring:
            # Two different source rows may contain intersecting weekday
            # patterns. Detect each duplicate real appointment individually.
            for weekday in parse_day_pattern(entry.day_pattern):
                token = (
                    entry.patient_id.casefold(), entry.treatment.casefold(),
                    weekday, entry.start_time,
                )
                if token in seen:
                    raise PrivateOutpatientSourceError(
                        "Outpatient recurring schedule contains overlapping duplicates"
                    )
                seen.add(token)
            if entry.treatment == "ΦΘ" and not entry.therapist_id:
                raise PrivateOutpatientSourceError(
                    "Outpatient physiotherapy requires an approved provider"
                )
            counts[entry.patient_id] += 1
            days += len(parse_day_pattern(entry.day_pattern))
        if any(count == 0 for count in counts.values()):
            raise PrivateOutpatientSourceError(
                "Outpatient registry contains a patient without approved appointments"
            )
        return PrivateOutpatientAudit(
            outpatient_patients=len(patients),
            recurring_treatments=len(recurring),
            weekday_placements=days,
            distinct_treatments=len({entry.treatment for entry in recurring}),
            source_unchanged=True,
        )
    except PrivateOutpatientSourceError:
        raise
    except Exception:
        raise PrivateOutpatientSourceError(
            "Outpatient appointments do not satisfy existing production rules"
        ) from None
    finally:
        if not source.is_file() or _hash(source) != before:
            raise PrivateOutpatientSourceError(
                "SAFETY STOP: private outpatient source changed during read-only audit"
            )
