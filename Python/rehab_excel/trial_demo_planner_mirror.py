"""Conservative visible PATIENT_PLANNER mirror of granular weekly appointments.

PROGRAM_WEEKLY_SLOTS stays authoritative for the daily engine. The legacy
MASTER/PATIENT_PLANNER view can only show one clock and provider per treatment;
display only losslessly representable rows. Complex cases remain visible in
DEMO_WEEKLY_VIEW instead of being collapsed into a misleading legacy slot.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import time
from typing import Iterable

from rehab_excel.weekly_slots_source import WEEKDAY_PATTERN

# Exact workbook column labels, unchanged from the original planning UI.
PLANNER_TREATMENT_HEADERS: dict[str, tuple[str, str, str | None]] = {
    "ΦΘ": ("ΦΘ_Ώρα", "ΦΘ_Ημέρες", "ΦΘ_Θεραπευτής"),
    "Ανακλινόμενο": ("Ανακλινόμενο_Ώρα", "Ανακλινόμενο_Ημέρες", None),
    "Πισίνα": ("Πισίνα_Ώρα", "Πισίνα_Ημέρες", None),
    "Ρομποτικό": ("Ρομποτικό_Ώρα", "Ρομποτικό_Ημέρες", "Ρομποτικό_Θεραπευτής"),
    "Εργο": ("Εργο_Ώρα", "Εργο_Ημέρες", "Εργο_Θεραπευτής"),
    "Λογο": ("Λογο_Ώρα", "Λογο_Ημέρες", "Λογο_Θεραπευτής"),
    "ΕΦΑ": ("ΕΦΑ_Ώρα", "ΕΦΑ_Ημέρες", None),
    "Ψυχ": ("Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος"),
}
WEEKDAY_ORDER = ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")


@dataclass(frozen=True)
class PlannerDemoMirror:
    patient_id: str
    treatment: str
    start_time: time
    day_pattern: str
    therapist: str | None


def choose_planner_demo_mirrors(
    weekly_rows: Iterable[tuple[str, str, str, time, str | None, str]]
) -> tuple[PlannerDemoMirror, ...]:
    """Mirrors one treatment only if all its weekday rows share time/provider.

    Exact specialty/weekdays are preserved. Never interpret a same-day
    double appointment as one cell. Distinct clocks or providers go to the
    granular workbook view alone.
    """
    groups = defaultdict(list)
    for patient_id, treatment, weekday, clock, provider, source_cell in weekly_rows:
        if weekday not in WEEKDAY_PATTERN:
            raise ValueError("Unsupported practice-demo weekday")
        if not isinstance(clock, time):
            raise ValueError("Practice-demo weekly clock must be a time")
        groups[(patient_id, treatment)].append((weekday, clock, provider))
    mirrors = []
    for (patient_id, treatment), entries in groups.items():
        if treatment not in PLANNER_TREATMENT_HEADERS:
            continue
        days = {item[0] for item in entries}
        if len(days) != len(entries):
            continue
        clocks = {item[1] for item in entries}
        providers = {item[2] for item in entries}
        if len(clocks) != 1 or len(providers) != 1:
            continue
        pattern = (
            "Καθ/να" if len(days) == len(WEEKDAY_ORDER)
            else "-".join(WEEKDAY_PATTERN[day] for day in WEEKDAY_ORDER if day in days)
        )
        mirrors.append(
            PlannerDemoMirror(
                patient_id, treatment, next(iter(clocks)), pattern,
                next(iter(providers)),
            )
        )
    return tuple(mirrors)
