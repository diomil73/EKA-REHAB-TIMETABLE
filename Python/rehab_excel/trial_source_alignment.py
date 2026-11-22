"""Read-only provenance audit between the original A:V programme and review XLSX.

This is a source-alignment check, not identity approval, admission, or import.
It never prints or returns patient names, IDs, room/doctor values, or source paths.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import re
import unicodedata

from openpyxl import load_workbook
from rehab_core.programme_mapping import ProgrammeMappingError, parse_clock


PROGRAMME_SHEET = "ΠΡΟΓΡΑΜΜΑ"
PATIENT_SHEET = "ΑΣΘΕΝΕΙΣ"
WEEKLY_SHEET = "ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"

PROGRAMME_HEADERS = (
    "ΩΡΑ", "ΗΜΕΡΑ", "ΦΥΣΙΚΟΘΕΡΑΠΕΙΑ", "ΑΝΑΚ/ΜΕΝΟ",
    "ΩΡΑ Ε/Θ", "ΗΜΕΡΑ Ε/Θ", "ΘΕΡΑΠΕΥΤΡΙΑ Ε/Θ",
    "ΩΡΑ Λ/Θ", "ΗΜΕΡΕΣ Λ/Θ", "ΘΕΡΑΠΕΥΤΡΙΑ Λ/Θ",
    "ΩΡΑ Ψ/Λ", "ΗΜΕΡΕΣ Ψ/Λ", "ΘΕΡΑΠΕΥΤΡΙΑ Ψ/Λ",
    "ΩΡΕΣ ΕΦΑ", "ΜΕΡΑ ΕΦΑ", "ΩΡΑ ΡΟΜΠ", "ΜΕΡΕΣ ΡΟΜΠ",
    "ΘΕΡΑΠΕΥΤΗΣ", "ΑΠΟΓΕΥΜΑ", "ΠΙΣΙΝΑ",
)
PATIENT_HEADERS = (
    "Γραμμή πηγής", "Κλινική", "Ασθενής (πηγή)",
    "Θάλαμος/ιατρός (πηγή)", "Αντιστοίχιση PatientID", "Επιβεβαίωση",
)
WEEKLY_HEADERS = (
    "Κλινική", "Γραμμή", "Ασθενής", "Ειδικότητα", "Ημέρα", "Ώρα",
    "Θεραπευτής ΦΘ/Ρομποτ.", "Προαιρετικό όνομα", "Κελιά πηγής",
    "Αρχικό κείμενο", "Κατάσταση",
)

# Column letters are the operator-supplied A:V source contract.
TREATMENT_SOURCE_COLUMNS = {
    "ΦΘ": frozenset({"C", "D", "E"}),
    "ΑΝΑΚΛΙΝΟΜΕΝΟ": frozenset({"F"}),
    "ΕΡΓΟΘΕΡΑΠΕΙΑ": frozenset({"G", "H", "I"}),
    "ΛΟΓΟΘΕΡΑΠΕΙΑ": frozenset({"J", "K", "L"}),
    "ΨΥΧΟΛΟΓΙΑ": frozenset({"M", "N", "O"}),
    "ΕΦΑ": frozenset({"P", "Q"}),
    "ΡΟΜΠΟΤΙΚΗ": frozenset({"R", "S", "T"}),
    "ΠΙΣΙΝΑ": frozenset({"V"}),
}
_CELL_RE = re.compile(r"([A-Z]{1,3})([1-9][0-9]*)")


class TrialSourceAlignmentError(ValueError):
    """The original programme and its review copy cannot be aligned safely."""


@dataclass(frozen=True)
class TrialSourceAlignmentResult:
    patient_rows_checked: int
    weekly_rows_checked: int
    source_references_checked: int
    missing_patient_context_rows: int
    identity_mismatches: int
    schedule_identity_mismatches: int
    source_reference_mismatches: int
    unsupported_specialties: int
    blockers: tuple[str, ...]
    inputs_unchanged: bool

    @property
    def aligned(self) -> bool:
        return self.inputs_unchanged and not self.blockers


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _clean(value: object) -> str:
    return "" if value is None else str(value).strip()


def _normal(value: object) -> str:
    raw = unicodedata.normalize("NFD", _clean(value).upper())
    return "".join(ch for ch in raw if unicodedata.category(ch) != "Mn")


def _row_number(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    try:
        number = int(value)
    except (TypeError, ValueError, OverflowError):
        return None
    try:
        if float(value) != number:
            return None
    except (TypeError, ValueError, OverflowError):
        return None
    return number if number > 1 else None


def _headers_match(ws, expected: tuple[str, ...], *, start_col: int = 1) -> bool:
    observed = tuple(
        _normal(ws.cell(1, start_col + offset).value)
        for offset in range(len(expected))
    )
    wanted = tuple(_normal(value) for value in expected)
    return observed == wanted


def audit_trial_source_alignment(
    *, programme: str | Path, review: str | Path
) -> TrialSourceAlignmentResult:
    """Compare source-row identity and weekly source-cell provenance, read-only.

    A positive result means only that the review rows point back to the stated
    source rows/cells. It does not approve IDs, weekdays, times, providers,
    holds, overlaps, clinical semantics, or an Excel import.
    """
    source = Path(programme).resolve()
    staged = Path(review).resolve()
    if source == staged or source.suffix.casefold() != ".xlsx" or staged.suffix.casefold() != ".xlsx":
        raise TrialSourceAlignmentError("Two distinct .xlsx source and review inputs are required")
    if not source.is_file() or not staged.is_file():
        raise TrialSourceAlignmentError("One or more source inputs are missing")

    before = (_digest(source), _digest(staged))
    patient_checked = weekly_checked = refs_checked = missing_context = 0
    identity_mismatches = schedule_identity_mismatches = ref_mismatches = unsupported = 0
    blockers: list[str] = []
    try:
        src_wb = load_workbook(source, read_only=True, data_only=True)
        rev_wb = load_workbook(staged, read_only=True, data_only=True)
        try:
            if PROGRAMME_SHEET not in src_wb.sheetnames:
                raise TrialSourceAlignmentError("Original programme sheet is missing")
            for required in (PATIENT_SHEET, WEEKLY_SHEET):
                if required not in rev_wb.sheetnames:
                    raise TrialSourceAlignmentError(f"Review sheet is missing: {required}")

            src = src_wb[PROGRAMME_SHEET]
            patients = rev_wb[PATIENT_SHEET]
            weekly = rev_wb[WEEKLY_SHEET]
            if not _headers_match(src, PROGRAMME_HEADERS, start_col=3):
                raise TrialSourceAlignmentError("Original A:V programme headers differ from contract")
            if not _headers_match(patients, PATIENT_HEADERS):
                raise TrialSourceAlignmentError("Patient review headers differ from contract")
            if not _headers_match(weekly, WEEKLY_HEADERS):
                raise TrialSourceAlignmentError("Weekly review headers differ from contract")

            source_patients: dict[int, tuple[str, str]] = {}
            seen_source_rows: set[int] = set()
            for values in patients.iter_rows(min_row=2, max_col=6, values_only=True):
                if not any(_clean(value) for value in values):
                    continue
                patient_checked += 1
                source_row = _row_number(values[0])
                clinic, name, context = _clean(values[1]), _clean(values[2]), _clean(values[3])
                if source_row is None or not clinic or not name:
                    identity_mismatches += 1
                    continue
                if source_row in seen_source_rows:
                    identity_mismatches += 1
                    continue
                seen_source_rows.add(source_row)
                source_name = _clean(src.cell(source_row, 2).value)
                source_context = _clean(src.cell(source_row, 1).value)
                if not source_name or source_name != name or source_context != context:
                    identity_mismatches += 1
                    continue
                if not context:
                    missing_context += 1
                source_patients[source_row] = (clinic, name)

            if identity_mismatches:
                blockers.append("PATIENT_SOURCE_ROW_OR_CONTEXT_MISMATCH")
            if not source_patients:
                blockers.append("NO_PATIENT_ROWS_ALIGNED")

            for values in weekly.iter_rows(min_row=2, max_col=11, values_only=True):
                if not any(_clean(value) for value in values):
                    continue
                weekly_checked += 1
                clinic, raw_row, name, specialty = (
                    _clean(values[0]), values[1], _clean(values[2]), _clean(values[3])
                )
                source_row = _row_number(raw_row)
                if source_row is None or source_row not in source_patients:
                    schedule_identity_mismatches += 1
                    continue
                patient_clinic, patient_name = source_patients[source_row]
                if clinic != patient_clinic or name != patient_name:
                    schedule_identity_mismatches += 1
                    continue

                allowed = TREATMENT_SOURCE_COLUMNS.get(specialty)
                if allowed is None:
                    unsupported += 1
                    continue
                cell_text = _clean(values[8]).upper()
                # R6 legacy shorthand is recognized ONLY by this independent
                # original-source audit, never as an unbound admission approval.
                # Both exact legacy spellings are complete specialty column
                # groups; no partial/mixed/malformed shorthand is accepted.
                legacy_groups = {
                    "ΕΡΓΟΘΕΡΑΠΕΙΑ": ("G/H/I", ("G", "H", "I")),
                    "ΨΥΧΟΛΟΓΙΑ": ("M/N/O", ("M", "N", "O")),
                }
                legacy = legacy_groups.get(specialty)
                if legacy is not None and cell_text == legacy[0]:
                    source_clock = src.cell(source_row, ord(legacy[1][0]) - 64).value
                    try:
                        clock_matches = parse_clock(values[5]) == parse_clock(source_clock)
                    except (ProgrammeMappingError, TypeError, ValueError):
                        clock_matches = False
                    if not clock_matches:
                        ref_mismatches += 1
                        continue
                    refs_checked += len(legacy[1])
                    continue
                refs = _CELL_RE.findall(cell_text)
                addresses = [f"{col}{row}" for col, row in refs]
                if not refs or len(addresses) != len(set(addresses)):
                    ref_mismatches += 1
                    continue
                if any(int(row) != source_row or col not in allowed for col, row in refs):
                    ref_mismatches += 1
                    continue
                time_col = {
                    "ΦΘ": "C", "ΑΝΑΚΛΙΝΟΜΕΝΟ": "F",
                    "ΕΡΓΟΘΕΡΑΠΕΙΑ": "G", "ΛΟΓΟΘΕΡΑΠΕΙΑ": "J",
                    "ΨΥΧΟΛΟΓΙΑ": "M", "ΕΦΑ": "P", "ΡΟΜΠΟΤΙΚΗ": "R", "ΠΙΣΙΝΑ": "V",
                }[specialty]
                if time_col not in {col for col, _ in refs}:
                    ref_mismatches += 1
                    continue
                time_value = src.cell(source_row, ord(time_col) - 64).value
                if not _clean(time_value):
                    ref_mismatches += 1
                    continue
                refs_checked += len(refs)

            if schedule_identity_mismatches:
                blockers.append("WEEKLY_SOURCE_IDENTITY_MISMATCH")
            if ref_mismatches:
                blockers.append("WEEKLY_SOURCE_CELL_REFERENCE_MISMATCH")
            if unsupported:
                blockers.append("UNSUPPORTED_WEEKLY_SPECIALTY")
        finally:
            src_wb.close()
            rev_wb.close()
    finally:
        unchanged = source.is_file() and staged.is_file() and before == (_digest(source), _digest(staged))
        if not unchanged:
            raise TrialSourceAlignmentError("SAFETY STOP: an input changed during read-only alignment")

    return TrialSourceAlignmentResult(
        patient_rows_checked=patient_checked,
        weekly_rows_checked=weekly_checked,
        source_references_checked=refs_checked,
        missing_patient_context_rows=missing_context,
        identity_mismatches=identity_mismatches,
        schedule_identity_mismatches=schedule_identity_mismatches,
        source_reference_mismatches=ref_mismatches,
        unsupported_specialties=unsupported,
        blockers=tuple(blockers),
        inputs_unchanged=unchanged,
    )
