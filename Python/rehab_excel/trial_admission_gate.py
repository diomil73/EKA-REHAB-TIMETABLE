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
            # Never accept a provider in two different patient appointments at
            # the same weekday/time. Distinct patients do not trigger the
            # patient-slot overlap guard, so this independent gate is essential.
            provider_slots: dict[tuple[str, str, object], str] = {}
            provider_collisions: set[tuple[str, str, object]] = set()
            for item in admitted:
                if not item.therapist_name:
                    continue  # Optional-provider specialty still blocks patient time.
                provider_key = item.therapist_name.casefold()
                key = (provider_key, item.weekday, item.start_time)
                previous_patient = provider_slots.get(key)
                if previous_patient is not None and previous_patient != item.patient_id:
                    provider_collisions.add(key)
                provider_slots[key] = item.patient_id
            if provider_collisions:
                blockers.append(
                    f"{len(provider_collisions)} provider weekday/time double-bookings"
                )
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
