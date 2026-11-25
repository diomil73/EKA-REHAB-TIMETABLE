"""Pure, non-authorizing DAILY_INPUT therapist assignment membership reconciliation.

This gate compares independently supplied preflight assignments against a
separately captured readback of exact immutable assignment IDs and members.
A THERAPIST_DAILY display/occupied-clock match is NOT evidence of these IDs.
Both inputs must be independently sourced; no workbook source reader, signer,
clinical approval, phase report, journal append or Apply path is implemented.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time

from rehab_core.recurring_conflicts import _provider_key
from .stable_recurring_ids import is_stable_assignment_id


class TherapistAssignmentMembershipError(ValueError):
    """An exact, source-bound therapist assignment comparison failed."""


@dataclass(frozen=True)
class TherapistAssignmentFact:
    assignment_id: str
    patient_id: str
    treatment: str
    provider: str
    service_date: date
    start_time: time


@dataclass(frozen=True)
class TherapistAssignmentMembershipReadback:
    service_date: date
    compared_assignments: int
    distinct_provider_clocks: int
    approved_shared_clocks: int


def _facts(name: str, rows: tuple[TherapistAssignmentFact, ...], day: date):
    if type(rows) is not tuple or not rows:
        raise TherapistAssignmentMembershipError(f"Nonempty immutable {name} required")
    by_id = {}
    provider_slots = {}
    patient_slots = set()
    for item in rows:
        if type(item) is not TherapistAssignmentFact:
            raise TherapistAssignmentMembershipError(f"Invalid {name} entry")
        if (type(item.assignment_id) is not str
                or not is_stable_assignment_id(item.assignment_id)
                or type(item.patient_id) is not str or not item.patient_id.strip()
                or item.patient_id != item.patient_id.strip()
                or type(item.treatment) is not str or not item.treatment.strip()
                or type(item.provider) is not str or not item.provider.strip()
                or type(item.service_date) is not date or item.service_date != day
                or type(item.start_time) is not time
                or item.start_time.second or item.start_time.microsecond):
            raise TherapistAssignmentMembershipError(f"Malformed or cross-date {name} entry")
        provider = _provider_key(item.provider)
        if not provider or item.assignment_id in by_id:
            raise TherapistAssignmentMembershipError(f"Duplicate/invalid {name} assignment identity")
        patient_clock = (item.patient_id.casefold(), item.start_time)
        if patient_clock in patient_slots:
            raise TherapistAssignmentMembershipError(f"Duplicate patient time in {name}")
        patient_slots.add(patient_clock)
        identity = (
            item.assignment_id, item.patient_id, item.treatment,
            provider, item.service_date, item.start_time,
        )
        by_id[item.assignment_id] = identity
        provider_slots.setdefault((provider, item.start_time), set()).add(item.patient_id)
    return by_id, {key: frozenset(members) for key, members in provider_slots.items()}


def verify_therapist_assignment_memberships(
    *,
    service_date: date,
    expected_assignments: tuple[TherapistAssignmentFact, ...],
    observed_assignments: tuple[TherapistAssignmentFact, ...],
    approved_shared_memberships: tuple[tuple[str, time, frozenset[str]], ...],
) -> TherapistAssignmentMembershipReadback:
    """Require full IDs, patient/provider/treatment/date/clock and exact sharing.

    Observed facts must originate from an independent verifier of persisted
    saved assignment IDs and member IDs, NOT the rendered therapist labels.
    Approvals must come from earlier source-bound operator decisions, not a
    reconstructed list of colliding assignments. This has NO authorization use.
    """
    if type(service_date) is not date:
        raise TherapistAssignmentMembershipError("Explicit service date required")
    expected, slots = _facts("preflight", expected_assignments, service_date)
    actual, actual_slots = _facts("saved readback", observed_assignments, service_date)
    if expected != actual or slots != actual_slots:
        raise TherapistAssignmentMembershipError("Assignment membership differs from preflight")
    if type(approved_shared_memberships) is not tuple:
        raise TherapistAssignmentMembershipError("Immutable sharing approvals required")
    approved = {}
    for item in approved_shared_memberships:
        if (type(item) is not tuple or len(item) != 3
                or type(item[0]) is not str or not _provider_key(item[0])
                or type(item[1]) is not time or item[1].second or item[1].microsecond
                or type(item[2]) is not frozenset or len(item[2]) < 2
                or any(type(pid) is not str or not pid.strip() or pid != pid.strip()
                       for pid in item[2])
                or len({pid.casefold() for pid in item[2]}) != len(item[2])):
            raise TherapistAssignmentMembershipError("Malformed shared-slot approval")
        key = (_provider_key(item[0]), item[1])
        if key in approved:
            raise TherapistAssignmentMembershipError("Duplicate shared-slot approval")
        approved[key] = item[2]
    required = {key: members for key, members in slots.items() if len(members) > 1}
    if required != approved:
        raise TherapistAssignmentMembershipError("Exact shared-slot membership approval differs")
    return TherapistAssignmentMembershipReadback(
        service_date=service_date,
        compared_assignments=len(expected),
        distinct_provider_clocks=len(slots),
        approved_shared_clocks=len(required),
    )
