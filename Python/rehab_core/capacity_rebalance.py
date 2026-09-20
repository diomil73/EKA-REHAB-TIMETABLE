from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .models import Session, Therapist


@dataclass(frozen=True)
class RebalanceSlot:
    provider_id: str
    target_date: date
    start_time: time
    sessions: tuple[Session, ...]

    @property
    def patient_ids(self) -> tuple[str, ...]:
        return tuple(session.patient_id for session in self.sessions)


@dataclass(frozen=True)
class ProviderCapacityCase:
    provider_id: str
    target_date: date
    used_timeslots: int
    max_timeslots: int
    excess_timeslots: int
    slots: tuple[RebalanceSlot, ...]


def build_capacity_rebalance_cases(
    *,
    target_date: date,
    sessions: Iterable[Session],
    therapists: Iterable[Therapist],
) -> tuple[ProviderCapacityCase, ...]:
    """Return providers whose *existing* dated programme exceeds capacity.

    Capacity is counted as distinct clock times. If two patients are active in
    the same provider/time slot on the same date, that is one occupied timeslot.
    To actually free such a slot during rebalancing, every active session in
    that slot must be moved, so the case keeps the whole slot group together.
    """

    therapist_map = {t.therapist_id: t for t in therapists}
    grouped: dict[str, dict[time, list[Session]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for session in sessions:
        if session.session_date != target_date:
            continue
        if session.therapist_id not in therapist_map:
            continue
        grouped[session.therapist_id][session.start_time].append(session)

    cases: list[ProviderCapacityCase] = []
    for provider_id, by_time in grouped.items():
        therapist = therapist_map[provider_id]
        used = len(by_time)
        limit = therapist.max_daily_timeslots
        if used <= limit:
            continue
        slots = tuple(
            RebalanceSlot(
                provider_id=provider_id,
                target_date=target_date,
                start_time=slot_time,
                sessions=tuple(
                    sorted(by_time[slot_time], key=lambda s: (s.patient_id, s.session_id))
                ),
            )
            for slot_time in sorted(by_time)
        )
        cases.append(
            ProviderCapacityCase(
                provider_id=provider_id,
                target_date=target_date,
                used_timeslots=used,
                max_timeslots=limit,
                excess_timeslots=used - limit,
                slots=slots,
            )
        )

    return tuple(sorted(cases, key=lambda case: case.provider_id.casefold()))
