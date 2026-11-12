"""Fail-closed admission gate for the isolated real-data Diomil trial.

Consumes the local-only reviewed staging workbook, NEVER uploads identities,
writes an .xlsm, assigns IDs by guessing, or silently drops unresolved cells.
A positive gate means the workbook is eligible for a separately audited
native-Windows Excel trial import; it is NOT evidence of an import.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook

from rehab_core.programme_mapping import (
    ProgrammeMappingError,
    SourceWeeklyEntry,
    WEEKDAYS,
    parse_clock,
    validate_weekly_entry_consistency,
)

SOURCE_PATIENT_SHEET = "ΑΣΘΕΝΕΙΣ"
SOURCE_WEEKLY_SHEET = "ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"
SOURCE_HOLDS_SHEET = "ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ"
SOURCE_OVERLAP_SHEET = "ΕΠΙΚΑΛΥΨΕΙΣ"
SOURCE_SHARED_SHEET = "ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS"
SHARED_HEADERS = ("Therapist", "Weekday", "StartTime", "PatientIDs", "Decision")
SHARED_ACCEPTED = "APPROVED_SHARED"
REQUIRED_SHEETS = frozenset({
    SOURCE_PATIENT_SHEET, SOURCE_WEEKLY_SHEET,
    SOURCE_HOLDS_SHEET, SOURCE_OVERLAP_SHEET,
})
CONFIRMED_MARKERS = frozenset({"CONFIRMED", "ΕΠΙΒΕΒΑΙΩΜΕΝΟ", "ΕΠΙΒΕΒΑΙΩΘΗΚΕ"})
SOURCE_TREATMENTS = {
    "ΦΘ": "ΦΘ",
    "ΡΟΜΠΟΤΙΚΗ": "Ρομποτικό",
    "ΕΡΓΟΘΕΡΑΠΕΙΑ": "Εργο",
    "ΛΟΓΟΘΕΡΑΠΕΙΑ": "Λογο",
    "ΨΥΧΟΛΟΓΙΑ": "Ψυχ",
    "ΕΦΑ": "ΕΦΑ",
    "ΑΝΑΚΛΙΝΟΜΕΝΟ": "Ανακλινόμενο",
    "ΠΙΣΙΝΑ": "Πισίνα",
}
SRC_CELL_RE = re.compile(r"\b([A-Z]{1,3}[1-9][0-9]*)\b")


class TrialAdmissionError(ValueError):
    """Source workbook incomplete or not independently reconciled."""


@dataclass(frozen=True)
class TrialAdmissionResult:
    candidate_patients: int
    verified_patient_ids: int
    weekly_proposals: int
    unadmitted_weekly_rows: int
    unresolved_source_rows: int
    overlap_rows: int
    duplicate_ids: int
    admitted_entries: tuple[SourceWeeklyEntry, ...]
    blockers: tuple[str, ...]

    @property
    def ready(self) -> bool:
        return not self.blockers


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _data_rows(ws, width: int):
    for number, values in enumerate(
        ws.iter_rows(min_row=2, max_col=width, values_only=True), start=2
    ):
        if any(_text(value) for value in values):
            yield number, values


def evaluate_trial_admission(path: str | Path) -> TrialAdmissionResult:
    """Perform full review, returning aggregate blockers without clinical names.

    The exact source row and hospital ward are used locally to join identities;
    names are cross-checked but never printed or included in the response.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        missing = REQUIRED_SHEETS.difference(wb.sheetnames)
        if missing:
            raise TrialAdmissionError("Missing review sheets: " + ", ".join(sorted(missing)))
        patients = tuple(_data_rows(wb[SOURCE_PATIENT_SHEET], 6))
        proposals = tuple(_data_rows(wb[SOURCE_WEEKLY_SHEET], 11))
        holds = tuple(_data_rows(wb[SOURCE_HOLDS_SHEET], 7))
        overlaps = tuple(_data_rows(wb[SOURCE_OVERLAP_SHEET], 7))

        patient_keys: dict[tuple[str, str, str], str] = {}
        verified_ids: list[str] = []
        seen_patient_rows: set[tuple[str, str]] = set()
        blockers: list[str] = []
        for excel_row, row in patients:
            source_row, clinic, name, _ward, patient_id, confirmation = row
            key = (_text(clinic), _text(source_row), _text(name))
            location = key[:2]
            if not all(key):
                blockers.append(f"Patient staging row {excel_row}: missing source identity")
                continue
            if location in seen_patient_rows:
                blockers.append(f"Patient staging row {excel_row}: duplicate clinic/source row")
            seen_patient_rows.add(location)
            if key in patient_keys:
                blockers.append(f"Patient staging row {excel_row}: repeated source identity")
            patient_keys[key] = _text(patient_id)
            if not _text(patient_id) or _text(confirmation).upper() not in CONFIRMED_MARKERS:
                continue
            verified_ids.append(_text(patient_id))

        duplicate_ids = len(verified_ids) - len({x.casefold() for x in verified_ids})
        if duplicate_ids:
            blockers.append(f"{duplicate_ids} duplicate confirmed PatientIDs")
        if len(verified_ids) != len(patients):
            blockers.append(
                f"{len(patients) - len(verified_ids)} patient IDs lack explicit confirmation"
            )
        if holds:
            blockers.append(f"{len(holds)} unresolved source cells")
        if overlaps:
            blockers.append(f"{len(overlaps)} cross-specialty overlap review rows")

        admitted: list[SourceWeeklyEntry] = []
        unadmitted = 0
        for excel_row, row in proposals:
            clinic, source_row, name, specialty, day, clock, required_name, optional_name, cells, _raw, state = row
            key = (_text(clinic), _text(source_row), _text(name))
            patient_id = patient_keys.get(key)
            if not patient_id:
                unadmitted += 1
                continue
            if _text(state) != "PROPOSED_EXACT_SOURCE":
                unadmitted += 1
                continue
            treatment = SOURCE_TREATMENTS.get(_text(specialty))
            match = SRC_CELL_RE.search(_text(cells))
            if treatment is None or _text(day) not in WEEKDAYS or match is None:
                blockers.append(f"Weekly staging row {excel_row}: unsupported specialty, day or source cell")
                unadmitted += 1
                continue
            provider = _text(required_name) if treatment in {"ΦΘ", "Ρομποτικό"} else _text(optional_name)
            if treatment in {"ΦΘ", "Ρομποτικό"} and not provider:
                blockers.append(f"Weekly staging row {excel_row}: mandatory provider missing")
                unadmitted += 1
                continue
            try:
                admitted.append(SourceWeeklyEntry(
                    patient_id=patient_id,
                    treatment=treatment,
                    weekday=_text(day),
                    start_time=parse_clock(clock),
                    therapist_name=provider or None,
                    source_cell=match.group(1),
                ))
            except (ProgrammeMappingError, ValueError):
                blockers.append(f"Weekly staging row {excel_row}: malformed weekday slot")
                unadmitted += 1

        if unadmitted:
            blockers.append(f"{unadmitted} weekly proposals cannot be admitted")
        if admitted:
            # Explicit *exact membership* approval is required for shared
            # therapist slots. Approved patient IDs are never stored in GitHub,
            # and a partial/stale approval never admits additional patients.
            actual_groups: dict[tuple[str, str, object], set[str]] = {}
            names_by_key: dict[tuple[str, str, object], str] = {}
            for item in admitted:
                if item.therapist_name:
                    key = (item.therapist_name.casefold(), item.weekday, item.start_time)
                    actual_groups.setdefault(key, set()).add(item.patient_id)
                    names_by_key[key] = item.therapist_name

            approved: dict[tuple[str, str, object], frozenset[str]] = {}
            if SOURCE_SHARED_SHEET in wb.sheetnames:
                sheet = wb[SOURCE_SHARED_SHEET]
                headers = tuple(sheet.cell(1, col).value for col in range(1, 6))
                if headers != SHARED_HEADERS:
                    blockers.append("Shared-slot approval sheet headers are invalid")
                else:
                    for approval_row, values in _data_rows(sheet, 5):
                        provider, weekday, clock, ids, decision = values
                        try:
                            key = (_text(provider).casefold(), _text(weekday),
                                   parse_clock(clock))
                        except (ProgrammeMappingError, TypeError, ValueError):
                            blockers.append(
                                f"Shared-slot approval row {approval_row}: invalid time"
                            )
                            continue
                        member_ids = [_text(x) for x in _text(ids).split("|")]
                        if (
                            not key[0] or key[1] not in WEEKDAYS
                            or _text(decision) != SHARED_ACCEPTED
                            or len(member_ids) < 2
                            or any(not x for x in member_ids)
                            or len(set(member_ids)) != len(member_ids)
                            or key in approved
                        ):
                            blockers.append(
                                f"Shared-slot approval row {approval_row}: invalid or duplicated approval"
                            )
                            continue
                        approved[key] = frozenset(member_ids)
            collisions = {
                key: frozenset(members)
                for key, members in actual_groups.items() if len(members) > 1
            }
            unapproved = sum(
                1 for key, members in collisions.items()
                if approved.get(key) != members
            )
            stale = sum(
                1 for key, members in approved.items()
                if collisions.get(key) != members
            )
            if unapproved:
                blockers.append(f"{unapproved} unapproved shared provider weekday/time slots")
            if stale:
                blockers.append(f"{stale} stale or mismatched shared-slot approvals")
            try:
                validate_weekly_entry_consistency(admitted)
            except ProgrammeMappingError:
                blockers.append("Patient cross-specialty duplicate slot detected")
        if not patients or not proposals:
            blockers.append("No source patients or weekly programme rows")
        return TrialAdmissionResult(
            candidate_patients=len(patients),
            verified_patient_ids=len(verified_ids),
            weekly_proposals=len(proposals),
            unadmitted_weekly_rows=unadmitted,
            unresolved_source_rows=len(holds),
            overlap_rows=len(overlaps),
            duplicate_ids=duplicate_ids,
            admitted_entries=tuple(admitted) if not blockers else (),
            blockers=tuple(blockers),
        )
    finally:
        wb.close()
