"""Private snapshot-bound provider availability and shared-slot review.

Reuses the current source merge and distinct-day-clock capacity audit. No
guessing a provider's schedule or creating clinical approvals. This snapshot
is a separately authorized technical input, not a new operational roster.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import time
from pathlib import Path
import json
import re

from rehab_core.programme_mapping import WEEKDAYS
from rehab_core.recurring_conflicts import _provider_key
from rehab_excel.trial_demo_source_projection import merge_trial_demo_outpatients
from rehab_excel.trial_provider_capacity import audit_trial_provider_capacity

SCHEMA = "diomil-private-provider-schedule-v1"
SOURCE_KEYS = ("programme", "review", "shell", "outpatients")


class TrialProviderScheduleError(ValueError):
    """Private schedule snapshot was not sufficient for the bulk preflight."""


@dataclass(frozen=True)
class TrialProviderSchedule:
    source_digests: tuple[str, str, str, str] = field(repr=False)
    available_slots: tuple[tuple[str, str, frozenset[time]], ...] = field(repr=False)
    shared_slots: tuple[tuple[str, str, time, frozenset[str]], ...] = field(repr=False)


def _unique_pairs(items):
    result = {}
    for key, value in items:
        if key in result:
            raise TrialProviderScheduleError("Duplicate snapshot field")
        result[key] = value
    return result


def _clock(value):
    if not isinstance(value, str) or not re.fullmatch(r"(?:[01][0-9]|2[0-3]):[0-5][0-9]", value):
        raise TrialProviderScheduleError("Invalid explicit clock")
    return time.fromisoformat(value)


def _provider_day(item):
    label, day = item["provider"], item["weekday"]
    if not isinstance(label, str) or not _provider_key(label) or day not in WEEKDAYS:
        raise TrialProviderScheduleError("Invalid provider or weekday")
    return _provider_key(label), day


def read_private_provider_schedule(path):
    """Parse a private JSON containing literal, explicitly authorized decisions."""
    path = Path(path).resolve()
    if path.suffix.casefold() != ".json" or not path.is_file() or any(
        (parent / ".git").exists() for parent in path.parents
    ):
        raise TrialProviderScheduleError("Private schedule JSON outside Git required")
    try:
        with path.open("r", encoding="utf-8") as stream:
            raw = json.load(stream, object_pairs_hook=_unique_pairs)
        if (
            not isinstance(raw, dict)
            or set(raw) != {"schema", "sources", "availability", "shared_slots"}
            or raw["schema"] != SCHEMA or not isinstance(raw["sources"], dict)
            or set(raw["sources"]) != set(SOURCE_KEYS)
            or not isinstance(raw["availability"], list)
            or not isinstance(raw["shared_slots"], list)
        ):
            raise TrialProviderScheduleError("Invalid private provider schedule schema")
        hashes = tuple(raw["sources"][key] for key in SOURCE_KEYS)
        if any(not isinstance(item, str) or not re.fullmatch(r"[a-f0-9]{64}", item) for item in hashes):
            raise TrialProviderScheduleError("Invalid source snapshot digests")
        availability = {}
        for item in raw["availability"]:
            if (
                not isinstance(item, dict) or set(item) != {"provider", "weekday", "times", "decision"}
                or item["decision"] != "APPROVED" or not isinstance(item["times"], list)
            ):
                raise TrialProviderScheduleError("Explicit availability decision required")
            key = _provider_day(item)
            times = frozenset(_clock(value) for value in item["times"])
            if key in availability or len(times) != len(item["times"]):
                raise TrialProviderScheduleError("Duplicate provider/day or clock")
            availability[key] = times
        shared = {}
        for item in raw["shared_slots"]:
            if (
                not isinstance(item, dict)
                or set(item) != {"provider", "weekday", "time", "patient_ids", "decision"}
                or item["decision"] != "APPROVED_SHARED"
                or not isinstance(item["patient_ids"], list)
                or len(item["patient_ids"]) < 2
                or any(not isinstance(pid, str) or not pid.strip() or pid.strip() != pid
                       for pid in item["patient_ids"])
                or len({pid.casefold() for pid in item["patient_ids"]}) != len(item["patient_ids"])
            ):
                raise TrialProviderScheduleError("Explicit shared membership decision required")
            key = (*_provider_day(item), _clock(item["time"]))
            if key in shared:
                raise TrialProviderScheduleError("Duplicate shared slot")
            shared[key] = frozenset(item["patient_ids"])
        return TrialProviderSchedule(
            source_digests=hashes,
            available_slots=tuple((*key, values) for key, values in availability.items()),
            shared_slots=tuple((*key, values) for key, values in shared.items()),
        )
    except TrialProviderScheduleError:
        raise
    except (OSError, ValueError, TypeError, KeyError, UnicodeError):
        raise TrialProviderScheduleError("Unable to parse approved private provider schedule") from None


def merge_with_approved_provider_schedule(inpatient, outpatient, *, plan, schedule):
    """Verify exact source-bound decisions before any DEMO seed or publication."""
    if not isinstance(schedule, TrialProviderSchedule) or schedule.source_digests != (
        plan.programme_sha256, plan.review_sha256, plan.intake_sha256, outpatient.source_sha256
    ):
        raise TrialProviderScheduleError("Provider schedule belongs to different sources")
    available = {(provider, day): clocks for provider, day, clocks in schedule.available_slots}
    shared = {(provider, day, clock): members for provider, day, clock, members in schedule.shared_slots}
    if len(available) != len(schedule.available_slots) or len(shared) != len(schedule.shared_slots):
        raise TrialProviderScheduleError("Duplicate schedule decisions")
    # The existing outpatient merger validates only shared groups touching an
    # outpatient. Independently confirmed INPATIENT-only groups must also have
    # exact membership, or they could reach native Excel without authorization.
    # Use the same operational provider identity normalization as both readers.
    all_provider_members: dict[tuple[str, str, time], set[str]] = {}
    outpatient_keys: set[tuple[str, str, time]] = set()
    for pid, _treatment, day, clock, provider, _source in inpatient.weekly_slots.rows:
        if provider:
            all_provider_members.setdefault(
                (_provider_key(provider), day, clock), set(),
            ).add(pid)
    for pid, _name, _treatment, day, clock, provider, _source in outpatient.weekday_view_rows:
        if provider:
            key = (_provider_key(provider), day, clock)
            outpatient_keys.add(key)
            all_provider_members.setdefault(key, set()).add(pid)

    exact_shared = {
        key: frozenset(members)
        for key, members in all_provider_members.items() if len(members) > 1
    }
    if shared != exact_shared:
        raise TrialProviderScheduleError(
            "Stale or mismatched explicit shared-slot approval: exact approved membership required for every multi-patient provider slot"
        )
    # Reuse the existing outpatient collision gate with just its own scope.
    # Passing inpatient-only approvals to it would be a false surplus approval.
    outpatient_shared = {
        key: members for key, members in shared.items() if key in outpatient_keys
    }
    combined = merge_trial_demo_outpatients(
        inpatient, outpatient, shared_approvals=outpatient_shared,
    )
    placements = [
        (day, clock, provider)
        for _pid, _treatment, day, clock, provider, _source in combined.weekly_slots.rows
    ] + [
        (day, clock, provider)
        for _pid, _name, _treatment, day, clock, provider, _source in outpatient.weekday_view_rows
    ]
    # Explicit available clock sets preserve the distinction between distinct
    # occupied times and patient lines. No capacity defaults or inference.
    if any(
        clock not in available.get((_provider_key(provider), day), frozenset())
        for day, clock, provider in placements if provider
    ):
        raise TrialProviderScheduleError("Named provider clock is not explicitly available")
    report = audit_trial_provider_capacity(
        combined, approved_limits={key: len(clocks) for key, clocks in available.items()}
    )
    if not report.named_provider_capacity_clear:
        raise TrialProviderScheduleError("Explicit named-provider capacity exceeded or missing")
    return combined
