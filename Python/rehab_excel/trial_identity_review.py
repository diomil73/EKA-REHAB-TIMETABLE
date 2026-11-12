"""Make a separate local identity-review xlsx, never auto-approve PatientIDs.

The latest PROGRAMMA staging contains source clinic/row/name in columns A:C
and reserved operator fields E:F. Suggestions go ONLY in new columns G:H.
Old CC patient IDs are compared read-only. Original workbooks are untouched.

Sensitive input/output must remain in an approved institution-controlled folder;
never commit generated review workbooks, raw names or patient IDs to GitHub.
"""
from __future__ import annotations

from collections import Counter
import hashlib
from pathlib import Path
import shutil
from typing import Iterable

from openpyxl import load_workbook

from rehab_core.models import PatientType
from .patient_registry_source import read_patient_registry
from .trial_identity_reconciliation import (
    OldPatientIdentity,
    PatientIdentity,
    suggest_identity_matches,
)
from .trial_admission_gate import SOURCE_PATIENT_SHEET


class TrialIdentityReviewError(RuntimeError):
    """Fail closed when inputs, output or review schema are unsafe."""


SUGGESTION_HEADERS = ("Προτεινόμενο παλιό PatientID (όχι έγκριση)", "Κατάσταση αντιστοίχισης")
EXPECTED_SOURCE_HEADERS = (
    "Γραμμή πηγής", "Κλινική", "Ασθενής (πηγή)",
    "Θάλαμος/ιατρός (πηγή)", "Αντιστοίχιση PatientID", "Επιβεβαίωση",
)


def create_local_identity_review(
    *,
    staging_path: str | Path,
    old_demo_path: str | Path,
    output_path: str | Path,
) -> dict[str, int]:
    """Create a *new* local XLSX review copy, explicitly without confirmation.

    This is an identity candidate report, NOT an admitted roster and NOT an
    import authorization. Protect ID reuse: blank/unconfirmed E:F remain so.
    """
    source = Path(staging_path).resolve()
    old = Path(old_demo_path).resolve()
    output = Path(output_path).resolve()
    if source.suffix.lower() != ".xlsx" or old.suffix.lower() != ".xlsm":
        raise TrialIdentityReviewError("Expected staging .xlsx and old demo .xlsm")
    if output.suffix.lower() != ".xlsx":
        raise TrialIdentityReviewError("Output must be a new .xlsx")
    if source == output or old == output:
        raise TrialIdentityReviewError("Refusing to overwrite an input workbook")
    if output.exists():
        raise TrialIdentityReviewError("SAFETY STOP: output already exists")
    if not source.is_file() or not old.is_file():
        raise TrialIdentityReviewError("One or more input workbooks do not exist")

    def _sha256(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for piece in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(piece)
        return digest.hexdigest()

    source_digest = _sha256(source)
    old_digest = _sha256(old)
    old_people = tuple(
        OldPatientIdentity(person.patient_id, person.display_name)
        for person in read_patient_registry(old)
        if person.patient_type == PatientType.INPATIENT
    )
    source_book = load_workbook(source, read_only=True, data_only=True)
    try:
        if SOURCE_PATIENT_SHEET not in source_book.sheetnames:
            raise TrialIdentityReviewError("Staging has no patient identity review sheet")
        ws = source_book[SOURCE_PATIENT_SHEET]
        header = tuple(ws.cell(1, col).value for col in range(1, 7))
        if header != EXPECTED_SOURCE_HEADERS:
            raise TrialIdentityReviewError("Staging patient source headers changed")
        if any(ws.cell(1, col).value not in (None, "")
               for col in (7, 8)):
            raise TrialIdentityReviewError("Staging identity suggestion columns already used")
        patients: list[PatientIdentity] = []
        for values in ws.iter_rows(min_row=2, max_col=6, values_only=True):
            source_row, clinic, name, _ward, _assigned, _confirmation = values
            if all(value is None or str(value).strip() == "" for value in values):
                continue
            try:
                row_number = int(source_row)
            except (ValueError, TypeError) as exc:
                raise TrialIdentityReviewError("Invalid source patient row number") from exc
            patients.append(PatientIdentity(row_number, str(clinic or "").strip(),
                                           str(name or "").strip()))
    finally:
        source_book.close()

    suggestions = suggest_identity_matches(patients, old_people)
    # Copy-first, then decorate only the new copy. E:F are never touched.
    output.parent.mkdir(parents=True, exist_ok=True)
    created_by_this_call = False
    try:
        with output.open("xb") as sink:
            created_by_this_call = True
            with source.open("rb") as original:
                shutil.copyfileobj(original, sink)
        if _sha256(source) != source_digest or _sha256(old) != old_digest:
            raise TrialIdentityReviewError("One of the identity sources changed during review")
        if _sha256(output) != source_digest:
            raise TrialIdentityReviewError("Review output does not match the input source")
        wb = load_workbook(output)
        try:
            ws = wb[SOURCE_PATIENT_SHEET]
            ws.cell(1, 7).value, ws.cell(1, 8).value = SUGGESTION_HEADERS
            count = 0
            for row in range(2, ws.max_row + 1):
                if not any(ws.cell(row, col).value not in (None, "")
                           for col in range(1, 7)):
                    continue
                if count >= len(suggestions):
                    raise TrialIdentityReviewError("Patient source changed during review")
                item = suggestions[count]
                if (str(ws.cell(row, 1).value).strip() != str(item.source_row)
                        or str(ws.cell(row, 2).value or "").strip() != item.clinic):
                    raise TrialIdentityReviewError("Source row changed during review")
                ws.cell(row, 7).value = item.suggested_patient_id
                ws.cell(row, 8).value = item.state
                count += 1
            if count != len(suggestions):
                raise TrialIdentityReviewError("Patient count mismatch while writing review")
            wb.save(output)
        finally:
            wb.close()
        totals = Counter(x.state for x in suggestions)
        return {"source_rows": len(suggestions),
                "exact_unconfirmed_suggestions": totals["EXACT_NAME_CANDIDATE_UNCONFIRMED"],
                "requires_individual_review": len(suggestions) - totals["EXACT_NAME_CANDIDATE_UNCONFIRMED"]}
    except Exception:
        # An atomic 'xb' failure means another writer owns this path.
        # NEVER delete that other file on our way out.
        if created_by_this_call:
            output.unlink(missing_ok=True)
        raise
