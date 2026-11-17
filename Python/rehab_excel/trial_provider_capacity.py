"""Read-only bulk DEMO capacity diagnostic using approved explicit limits.

No default maximum, name-to-ID inference, capacity exception approval, or
native write. Callers must supply independently approved provider/day limits
for the particular source snapshot. Shared memberships remain separately
validated by the admission/merge gates. This is not a release gate on its own.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Mapping

from rehab_core.capacity import validate_recurring_timeslot_capacity
from rehab_core.models import BaseScheduleEntry
from rehab_core.programme_mapping import WEEKDAYS
from rehab_core.recurring_conflicts import _provider_key
from rehab_excel.trial_demo_source_projection import TrialDemoSourceProjection


class TrialProviderCapacityError(ValueError):
    """Malformed input, never includes source-identifying values."""


@dataclass(frozen=True)
class TrialProviderCapacityAudit:
    provider_days: int
    named_provider_placements: int
    occupied_distinct_slots: int
    missing_limit_days: int
    exceeded_limit_days: int
    unassigned_placements: int

    @property
    def named_provider_capacity_clear(self) -> bool:
        return not self.missing_limit_days and not self.exceeded_limit_days


def audit_trial_provider_capacity(
    projection: TrialDemoSourceProjection,
    *, approved_limits: Mapping[tuple[str, str], int],
) -> TrialProviderCapacityAudit:
    """Count only named provider clock slots, per weekday, across both sources.

    Limits are exact provider-label/weekday integers supplied independently
    (including zero). Reuse existing operational normalization and capacity
    primitive. Unknown capacity is not zero and never falls back to six.
    Blank optional providers remain unassigned, not attributed to a person.
    No interval/duration, service availability or daily overlay is inferred.
    """
    if not isinstance(projection, TrialDemoSourceProjection):
        raise TrialProviderCapacityError("Expected a source projection")
    limits = {}
    for key, limit in approved_limits.items():
        if (
            not isinstance(key, tuple) or len(key) != 2
            or not isinstance(key[0], str) or key[1] not in WEEKDAYS
            or not _provider_key(key[0]) or type(limit) is not int or limit < 0
        ):
            raise TrialProviderCapacityError("Invalid approved provider/day limit")
        normalized = (_provider_key(key[0]), key[1])
        if normalized in limits:
            raise TrialProviderCapacityError("Duplicate normalized approved provider/day limit")
        limits[normalized] = limit

    placements = [
        (pid, treatment, day, clock, provider)
        for pid, treatment, day, clock, provider, _ in projection.weekly_slots.rows
    ]
    if projection.outpatient is not None:
        placements.extend(
            (pid, treatment, day, clock, provider)
            for pid, _name, treatment, day, clock, provider, _
            in projection.outpatient.weekday_view_rows
        )
    if len(placements) != projection.weekday_slot_count:
        raise TrialProviderCapacityError("Projection placement count differs")
    grouped = {}
    unassigned = 0
    named = 0
    for index, (pid, treatment, day, clock, provider) in enumerate(placements, 1):
        if day not in WEEKDAYS or not isinstance(clock, time):
            raise TrialProviderCapacityError("Invalid projected weekday/clock")
        if provider is None or provider == "":
            unassigned += 1
            continue
        if not isinstance(provider, str) or not _provider_key(provider):
            raise TrialProviderCapacityError("Invalid projected provider label")
        normalized = _provider_key(provider)
        named += 1
        entry = BaseScheduleEntry(
            base_entry_id=f"capacity-preview:{index}", patient_id=pid,
            treatment=treatment, start_time=clock, day_pattern=day,
            therapist_id=normalized, robotic=False,
        )
        grouped.setdefault((normalized, day), []).append(entry)
    missing = exceeded = used = 0
    for key, entries in grouped.items():
        clocks = {entry.start_time for entry in entries}
        used += len(clocks)
        if key not in limits:
            missing += 1
            continue
        limit = limits[key]
        # The existing recurring capacity primitive accepts limits >= 1.
        # Zero is an explicit unavailable day, not a missing or default limit.
        if limit == 0:
            exceeded += 1
            continue
        decision = validate_recurring_timeslot_capacity(
            provider_id=key[0], max_daily_timeslots=limit,
            existing_entries=entries, proposed_day_pattern=key[1],
            proposed_time=entries[0].start_time,
        )
        if not decision.allowed:
            exceeded += 1
    return TrialProviderCapacityAudit(
        provider_days=len(grouped), named_provider_placements=named,
        occupied_distinct_slots=used, missing_limit_days=missing,
        exceeded_limit_days=exceeded, unassigned_placements=unassigned,
    )
