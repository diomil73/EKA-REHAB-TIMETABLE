"""Effective-dated changes to an individual recurring therapy assignment.

This is separate from daily replacement/absence overrides. A permanent change
takes effect starting on a chosen date, and old dates continue to materialize
the original assignment for historical read-back. No Excel writes here.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date, time
from typing import Iterable

from .base_schedule import materialize_sessions_for_date
from .models import BaseScheduleEntry, Session, Therapist
from .permanent_assignment import check_permanent_assignment
from .day_patterns import parse_day_pattern


class RecurringTherapyChangeError(ValueError):
    """Unsafe permanent treatment-assignment transition."""


@dataclass(frozen=True)
class RecurringTherapyChange:
    base_entry_id: str
    patient_id: str
    effective_date: date
    new_therapist_id: str | None = None
    new_start_time: time | None = None
    new_day_pattern: str | None = None
    reason: str = ""

    def __post_init__(self):
        if not self.base_entry_id.strip() or not self.patient_id.strip():
            raise RecurringTherapyChangeError("Stable assignment and PatientID required")
        if (self.new_therapist_id is None and self.new_start_time is None
                and self.new_day_pattern is None):
            raise RecurringTherapyChangeError("At least one schedule field must change")
        if self.new_therapist_id is not None and not self.new_therapist_id.strip():
            raise RecurringTherapyChangeError("New provider cannot be blank")
        if self.new_day_pattern is not None:
            parse_day_pattern(self.new_day_pattern)


def recurring_entries_on_date(
    entries: Iterable[BaseScheduleEntry],
    *,
    baseline_date: date,
    target_date: date,
    changes: Iterable[RecurringTherapyChange] = (),
) -> tuple[BaseScheduleEntry, ...]:
    """Derive effective recurrence; original BaseScheduleEntry objects survive.

    Each per-day granular entry can be changed independently. Patient identity,
    treatment, and base entry ID remain fixed. Use separate additions/removals
    for changing the number of weekly sessions; a day-pattern change must
    still pass patient/provider capacity checks at operational approval time.
    """
    if target_date < baseline_date:
        raise RecurringTherapyChangeError("Cannot project before known baseline")
    original = tuple(entries)
    by_id = {item.base_entry_id: item for item in original}
    if len(by_id) != len(original):
        raise RecurringTherapyChangeError("Duplicate recurring assignment ID")
    grouped: dict[str, list[RecurringTherapyChange]] = {}
    change_keys: set[tuple[str, date]] = set()
    for change in changes:
        entry = by_id.get(change.base_entry_id)
        if entry is None or entry.patient_id != change.patient_id:
            raise RecurringTherapyChangeError("Unknown or mismatched recurring assignment")
        if change.effective_date < baseline_date:
            raise RecurringTherapyChangeError("Change predates known baseline")
        key = (change.base_entry_id, change.effective_date)
        if key in change_keys:
            raise RecurringTherapyChangeError("Multiple changes to assignment on same date")
        change_keys.add(key)
        grouped.setdefault(change.base_entry_id, []).append(change)
    effective = []
    for item in original:
        revised = item
        for change in sorted(
            grouped.get(item.base_entry_id, ()), key=lambda x: x.effective_date
        ):
            if change.effective_date > target_date:
                break
            revised = replace(
                revised,
                therapist_id=(revised.therapist_id if change.new_therapist_id is None
                              else change.new_therapist_id),
                start_time=(revised.start_time if change.new_start_time is None
                            else change.new_start_time),
                day_pattern=(revised.day_pattern if change.new_day_pattern is None
                             else change.new_day_pattern),
            )
        effective.append(revised)
    return tuple(effective)



def validate_effective_weekly_patient_slots(
    entries: Iterable[BaseScheduleEntry],
) -> None:
    """Reject same-patient weekday/time collisions in an effective programme.

    This deliberately checks EVERY weekday rather than only the day on which
    the operator edits a permanent therapy. Two different therapies, or two
    assignments of the same therapy, may not occupy a patient's same hour.
    Provider shared-slot approval and daily-capacity checks are separate.
    """
    occupied: set[tuple[str, object, time]] = set()
    for entry in entries:
        for weekday in parse_day_pattern(entry.day_pattern):
            key = (entry.patient_id, weekday, entry.start_time)
            if key in occupied:
                raise RecurringTherapyChangeError(
                    "Permanent change creates a weekday patient-slot collision"
                )
            occupied.add(key)


def materialize_effective_sessions(
    entries: Iterable[BaseScheduleEntry],
    target_date: date,
    *,
    baseline_date: date,
    changes: Iterable[RecurringTherapyChange] = (),
) -> list[Session]:
    """Use the same date for effective recurrence and session materialization."""
    return materialize_sessions_for_date(
        recurring_entries_on_date(
            entries, baseline_date=baseline_date,
            target_date=target_date, changes=changes,
        ),
        target_date,
    )


@dataclass(frozen=True)
class ValidatedRecurringChangePreview:
    """An authorized *proposal*, not a saved event or Excel modification."""

    change: RecurringTherapyChange
    effective_entries: tuple[BaseScheduleEntry, ...]
    affected_weekdays: tuple[str, ...]


def prepare_validated_recurring_change(
    entries: Iterable[BaseScheduleEntry],
    *,
    baseline_date: date,
    change: RecurringTherapyChange,
    current_changes: Iterable[RecurringTherapyChange] = (),
    therapists: Iterable[Therapist] = (),
) -> ValidatedRecurringChangePreview:
    """Preflight a permanent edit over every weekday before it can be approved.

    Existing effective events are projected as of the new event's start.
    Fail closed on duplicate/conflicting patient times and provider capacity or
    same-clock collisions. Known providers are mandatory for provider-owned
    treatments. Existing approved therapist groups need a separate explicit
    reassessment rather than silent acceptance in this strict edit workflow.
    """
    source = tuple(entries)
    history = tuple(current_changes)
    candidate_history = (*history, change)
    future = recurring_entries_on_date(
        source, baseline_date=baseline_date,
        target_date=change.effective_date, changes=candidate_history,
    )
    validate_effective_weekly_patient_slots(future)
    matching = [entry for entry in future if entry.base_entry_id == change.base_entry_id]
    if len(matching) != 1 or matching[0].patient_id != change.patient_id:
        raise RecurringTherapyChangeError("Changed recurring entry cannot be resolved")
    modified = matching[0]
    providers = {provider.therapist_id: provider for provider in therapists}
    if modified.therapist_id is not None:
        provider = providers.get(modified.therapist_id)
        if provider is None:
            raise RecurringTherapyChangeError("Changed assignment has unknown therapist")
        outcome = check_permanent_assignment(
            provider_id=provider.therapist_id,
            max_daily_timeslots=provider.max_daily_timeslots,
            existing_entries=future,
            proposed_day_pattern=modified.day_pattern,
            proposed_time=modified.start_time,
            source_entry_id=modified.base_entry_id,
            patient_id=modified.patient_id,
        )
        if not outcome.allowed:
            raise RecurringTherapyChangeError(
                "Permanent change has therapist collision, patient collision or capacity excess"
            )
    elif modified.treatment in {"ΦΘ", "Ρομποτικό"}:
        raise RecurringTherapyChangeError("Required therapist is missing")

    # A newly inserted event must remain safe when older, already scheduled
    # future events take effect. Checking only its start date is insufficient:
    # a patient or provider conflict could first appear next week/month.
    for future_date in sorted({
        event.effective_date for event in candidate_history
        if event.effective_date > change.effective_date
    }):
        later_entries = recurring_entries_on_date(
            source, baseline_date=baseline_date,
            target_date=future_date, changes=candidate_history,
        )
        validate_effective_weekly_patient_slots(later_entries)
        future_modified = next(
            item for item in later_entries
            if item.base_entry_id == change.base_entry_id
        )
        if future_modified.therapist_id is not None:
            future_provider = providers.get(future_modified.therapist_id)
            if future_provider is None:
                raise RecurringTherapyChangeError(
                    "Future schedule change references unknown therapist"
                )
            future_check = check_permanent_assignment(
                provider_id=future_provider.therapist_id,
                max_daily_timeslots=future_provider.max_daily_timeslots,
                existing_entries=later_entries,
                proposed_day_pattern=future_modified.day_pattern,
                proposed_time=future_modified.start_time,
                source_entry_id=future_modified.base_entry_id,
                patient_id=future_modified.patient_id,
            )
            if not future_check.allowed:
                raise RecurringTherapyChangeError(
                    "Future change boundary creates therapist collision or capacity excess"
                )
        elif future_modified.treatment in {"ΦΘ", "Ρομποτικό"}:
            raise RecurringTherapyChangeError(
                "Future change removes required therapist"
            )
    return ValidatedRecurringChangePreview(
        change=change,
        effective_entries=future,
        affected_weekdays=tuple(sorted(day.name for day in parse_day_pattern(
            modified.day_pattern
        ))),
    )
