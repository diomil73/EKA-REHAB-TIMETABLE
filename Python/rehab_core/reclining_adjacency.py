from __future__ import annotations

from collections import defaultdict
from datetime import datetime, time
from typing import Iterable
import unicodedata

from .base_schedule import pattern_applies_on_date
from .global_replacement_integration import AdjacencyPenalty
from .models import BaseScheduleEntry, Session


def _norm(value: str | None) -> str:
    decomposed = unicodedata.normalize("NFD", str(value or "").strip().casefold())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _minutes(value: time) -> int:
    return value.hour * 60 + value.minute


def build_reclining_adjacency_penalty(
    base_entries: Iterable[BaseScheduleEntry],
) -> AdjacencyPenalty:
    """Prefer replacement ΦΘ slots that stay close to reclining treatment.

    Reclining entries have no provider and therefore are intentionally absent from
    materialized ``Session`` objects.  The penalty is derived from the recurring
    base schedule for the target service date.  It is the minute distance from a
    candidate ΦΘ time to the nearest reclining treatment of the same patient.

    The global solver treats this as a lexicographic soft preference after maximum
    placement count and before original-time preservation.  Patients without a
    same-day reclining treatment, and non-ΦΘ target sessions, receive zero penalty.
    """

    reclining_by_patient: dict[str, tuple[BaseScheduleEntry, ...]] = {}
    grouped: dict[str, list[BaseScheduleEntry]] = defaultdict(list)
    for entry in base_entries:
        if _norm(entry.treatment) == _norm("Ανακλινόμενο"):
            grouped[entry.patient_id].append(entry)
    reclining_by_patient = {
        patient_id: tuple(entries) for patient_id, entries in grouped.items()
    }

    def penalty(target_session: Session, provider_id: str, replacement_time: time) -> int:
        del provider_id
        if _norm(target_session.treatment) != _norm("ΦΘ"):
            return 0

        same_day = tuple(
            entry
            for entry in reclining_by_patient.get(target_session.patient_id, ())
            if pattern_applies_on_date(entry.day_pattern, target_session.session_date)
        )
        if not same_day:
            return 0

        candidate_minutes = _minutes(replacement_time)
        return min(
            abs(candidate_minutes - _minutes(entry.start_time))
            for entry in same_day
        )

    return penalty
