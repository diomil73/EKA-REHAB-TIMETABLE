"""Pure patient-management draft contract for optional afternoon physiotherapy.

No workbook writes, patient registry mutation, guessed sessions, or source import.
The three fields must remain empty until an operator explicitly supplies all of
them. This is NOT a live UI/VBA integration.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from enum import Enum
from typing import Any

from rehab_core.programme_mapping import (
    ProgrammeMappingError, WEEKDAYS, parse_clock, parse_weekday_groups,
)


class PatientProgrammeDraftError(ValueError):
    """An incomplete or incorrectly scoped patient treatment draft."""


class TreatmentProgramme(str, Enum):
    MORNING = "MORNING"
    AFTERNOON = "AFTERNOON"


POOL_TREATMENT = "Πισίνα"
AFTERNOON_TREATMENT = "ΦΘ"


@dataclass(frozen=True)
class AfternoonPhysioDraft:
    """One optional weekly afternoon physiotherapy time/day/provider triple."""

    start_time: time | None = None
    weekday: str | None = None
    physiotherapist: str | None = None

    @property
    def blank(self) -> bool:
        return (
            self.start_time is None
            and self.weekday is None
            and self.physiotherapist is None
        )

    def __post_init__(self) -> None:
        if self.blank:
            return
        if (
            not isinstance(self.start_time, time)
            or self.weekday not in WEEKDAYS
            or not isinstance(self.physiotherapist, str)
            or not self.physiotherapist.strip()
        ):
            raise PatientProgrammeDraftError(
                "Afternoon physiotherapy requires time, explicit weekday and therapist together"
            )


def prepare_afternoon_physio_drafts(
    *, clock: Any = None, days: Any = None, therapist: Any = None
) -> tuple[AfternoonPhysioDraft, ...]:
    """Return no rows for a wholly empty form, without creating appointments.

    A single time/provider can cover multiple explicit days. Slash-separated
    time/provider groups are intentionally unsupported for this first slice;
    the operator may enter each day separately after the UI supports it.
    """
    values = (clock, days, therapist)
    if all(value is None or (isinstance(value, str) and not value.strip()) for value in values):
        return ()
    if any(value is None or (isinstance(value, str) and not value.strip()) for value in values):
        raise PatientProgrammeDraftError("Afternoon physiotherapy triple is incomplete")
    if not isinstance(therapist, str) or "/" in therapist or not therapist.strip():
        raise PatientProgrammeDraftError("Specify one named afternoon physiotherapist")
    if isinstance(clock, str) and "/" in clock:
        raise PatientProgrammeDraftError("Split different times into separate explicit entries")
    try:
        parsed_time = parse_clock(clock)
        groups = parse_weekday_groups(days)
    except (ProgrammeMappingError, TypeError, ValueError) as exc:
        raise PatientProgrammeDraftError("Invalid afternoon time or weekdays") from exc
    canonical_days = tuple(day for group in groups for day in group)
    return tuple(
        AfternoonPhysioDraft(parsed_time, day, therapist.strip())
        for day in canonical_days
    )


def validate_treatment_programme(treatment: str, programme: TreatmentProgramme) -> None:
    """Pool belongs to morning; afternoon only admits physiotherapy."""
    if not isinstance(programme, TreatmentProgramme) or not isinstance(treatment, str) or not treatment.strip():
        raise PatientProgrammeDraftError("Known treatment/programme required")
    if programme is TreatmentProgramme.AFTERNOON and treatment != AFTERNOON_TREATMENT:
        raise PatientProgrammeDraftError("Afternoon programme only accepts physiotherapy")
    if treatment == POOL_TREATMENT and programme is not TreatmentProgramme.MORNING:
        raise PatientProgrammeDraftError("Pool belongs to the morning programme")
