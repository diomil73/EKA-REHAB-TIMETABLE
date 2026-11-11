"""Read optional granular weekly slots for the existing PATIENT_PLANNER system.

One source row per weekday, time and specialty. This preserves multiple daily
placements for a specialty without changing the visual MASTER layout. The source
must be populated only after separate admission/census and conflict checks.

No workbook save and no patient names in this storage schema.
"""
from __future__ import annotations

from datetime import time
from typing import Iterable

from rehab_core.models import BaseScheduleEntry
from rehab_core.programme_mapping import (
    ProgrammeMappingError, SourceWeeklyEntry, WEEKDAYS,
    parse_clock, validate_weekly_entry_consistency,
)

SHEET_NAME = "PROGRAM_WEEKLY_SLOTS"
HEADERS = ("PatientID", "Treatment", "Weekday", "StartTime", "Therapist", "SourceCell")
REQUIRES_PROVIDER = frozenset({"ΦΘ", "Ρομποτικό"})
KNOWN_TREATMENTS = frozenset({"ΦΘ", "Ρομποτικό", "Ανακλινόμενο", "Πισίνα", "Εργο", "Λογο", "ΕΦΑ", "Ψυχ"})
WEEKDAY_PATTERN = {
    "ΔΕ": "Δε", "ΤΡ": "Τρ", "ΤΕ": "Τε", "ΠΕ": "Πε", "ΠΑ": "Πα",
}


class WeeklySlotsSourceError(ValueError):
    """The optional weekly source cannot be admitted safely."""


def read_weekly_slots_sheet(
    worksheet,
    *,
    inpatient_ids: Iterable[str],
) -> tuple[BaseScheduleEntry, ...]:
    """Validate every row before making ANY recurring entries available."""
    header = tuple(worksheet.cell(1, col).value for col in range(1, 7))
    if header != HEADERS:
        raise WeeklySlotsSourceError("PROGRAM_WEEKLY_SLOTS headers differ from contract")
    allowed = set(inpatient_ids)
    if not allowed:
        raise WeeklySlotsSourceError("No authoritative inpatient registry available")
    raw_entries: list[SourceWeeklyEntry] = []
    seen_addresses: set[tuple[str, str, str, str]] = set()
    for row_number, values in enumerate(
        worksheet.iter_rows(min_row=2, max_col=6, values_only=True), start=2
    ):
        if all(value is None or str(value).strip() == "" for value in values):
            continue
        patient_id, treatment, weekday, clock, therapist, source_cell = values
        patient_id = str(patient_id or "").strip()
        treatment = str(treatment or "").strip()
        weekday = str(weekday or "").strip().upper()
        therapist = str(therapist).strip() if therapist is not None else None
        source_cell = str(source_cell or "").strip().upper()
        if patient_id not in allowed:
            raise WeeklySlotsSourceError(f"Weekly slot row {row_number}: unknown or non-inpatient PatientID")
        if treatment not in KNOWN_TREATMENTS:
            raise WeeklySlotsSourceError(f"Weekly slot row {row_number}: unsupported treatment")
        if weekday not in WEEKDAYS:
            raise WeeklySlotsSourceError(f"Weekly slot row {row_number}: invalid weekday")
        if treatment in REQUIRES_PROVIDER and not therapist:
            raise WeeklySlotsSourceError(f"Weekly slot row {row_number}: therapist required")
        try:
            parsed_time = parse_clock(clock)
            entry = SourceWeeklyEntry(
                patient_id, treatment, weekday, parsed_time,
                therapist or None, source_cell,
            )
        except (ProgrammeMappingError, TypeError, ValueError) as exc:
            raise WeeklySlotsSourceError(
                f"Weekly slot row {row_number}: malformed source data"
            ) from exc
        duplicate_key = (patient_id, treatment, weekday, source_cell)
        if duplicate_key in seen_addresses:
            raise WeeklySlotsSourceError(f"Weekly slot row {row_number}: duplicate source entry")
        seen_addresses.add(duplicate_key)
        raw_entries.append(entry)
    try:
        validate_weekly_entry_consistency(raw_entries)
    except ProgrammeMappingError as exc:
        raise WeeklySlotsSourceError(str(exc)) from exc

    return tuple(
        BaseScheduleEntry(
            base_entry_id=f"weekly:{idx}",
            patient_id=item.patient_id,
            treatment=item.treatment,
            start_time=item.start_time,
            day_pattern=WEEKDAY_PATTERN[item.weekday],
            therapist_id=item.therapist_name,
            robotic=item.treatment == "Ρομποτικό",
        )
        for idx, item in enumerate(raw_entries, start=2)
    )


def overlay_weekly_entries(
    *,
    planner_entries: Iterable[BaseScheduleEntry],
    weekly_entries: Iterable[BaseScheduleEntry],
) -> list[BaseScheduleEntry]:
    """A granular (patient, treatment) pair replaces its legacy single planner cell.

    All other legacy entries remain as before. Never add the two schemas
    together for an overlapping pair, which would double book patients.
    """
    weekly = tuple(weekly_entries)
    replaced = {(item.patient_id, item.treatment) for item in weekly}
    return [
        *(item for item in planner_entries
          if (item.patient_id, item.treatment) not in replaced),
        *weekly,
    ]
