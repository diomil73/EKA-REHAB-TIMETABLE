"""Explicit membership review gate for shared afternoon physiotherapy slots.

This module never writes a workbook and never guesses group therapy approval.
Raw patient IDs are intentionally absent from error text and repr of reports.
Times are exact weekly start times; true overlapping intervals still require
duration-aware validation before native operational approval.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import time
from typing import Iterable

from rehab_excel.afternoon_slots_source import (
    AfternoonSlot, AfternoonSlotsError, validate_afternoon_rows,
)


class AfternoonSharedSlotError(ValueError):
    """An afternoon therapist's shared slot has no exact membership approval."""


@dataclass(frozen=True)
class SharedSlotApproval:
    therapist: str
    weekday: str
    start_time: time
    patient_ids: frozenset[str] = field(repr=False)

    def __post_init__(self) -> None:
        if (
            not isinstance(self.therapist, str) or not self.therapist.strip()
            or not isinstance(self.weekday, str) or not self.weekday.strip()
            or not isinstance(self.start_time, time)
            or not isinstance(self.patient_ids, frozenset)
            or len(self.patient_ids) < 2
            or any(not isinstance(pid, str) or not pid.strip() for pid in self.patient_ids)
        ):
            raise AfternoonSharedSlotError("Invalid explicit shared-slot approval")


@dataclass(frozen=True)
class SharedSlotGateReport:
    shared_slot_count: int
    matched_approval_count: int
    rejected_approval_count: int
    unapproved_slot_count: int

    @property
    def ready(self) -> bool:
        return not self.rejected_approval_count and not self.unapproved_slot_count


def review_afternoon_shared_slots(
    rows: Iterable[AfternoonSlot],
    approvals: Iterable[SharedSlotApproval] = (),
) -> SharedSlotGateReport:
    """Require *exact* member sets, with no stale, duplicate or partial consent.

    An isolated patient's own repeated time/day is rejected upstream.
    Different patients with the same therapist/day/start-time create a shared
    group requiring an exact approval. Approvals must not authorize a group
    that no longer exists. This is not an end-to-end clinical booking gate.
    """
    try:
        items = validate_afternoon_rows(rows)
    except AfternoonSlotsError as exc:
        raise AfternoonSharedSlotError("Invalid afternoon schedule") from exc
    groups: dict[tuple[str, str, time], set[str]] = {}
    for slot in items:
        key = (slot.therapist.strip().casefold(), slot.weekday, slot.start_time)
        groups.setdefault(key, set()).add(slot.patient_id)
    collisions = {
        key: frozenset(members)
        for key, members in groups.items() if len(members) > 1
    }
    accepted: dict[tuple[str, str, time], frozenset[str]] = {}
    bad = 0
    for approval in approvals:
        if not isinstance(approval, SharedSlotApproval):
            bad += 1
            continue
        key = (approval.therapist.strip().casefold(), approval.weekday, approval.start_time)
        if key in accepted:
            bad += 1
        else:
            accepted[key] = approval.patient_ids
    mismatched = sum(
        members != collisions.get(key)
        for key, members in accepted.items()
    )
    missing = sum(
        accepted.get(key) != members
        for key, members in collisions.items()
    )
    report = SharedSlotGateReport(
        shared_slot_count=len(collisions),
        matched_approval_count=sum(
            accepted.get(key) == members for key, members in collisions.items()
        ),
        rejected_approval_count=bad + mismatched,
        unapproved_slot_count=missing,
    )
    if not report.ready:
        raise AfternoonSharedSlotError(
            "SAFETY STOP: Missing, stale or duplicate shared-slot approval"
        )
    return report
