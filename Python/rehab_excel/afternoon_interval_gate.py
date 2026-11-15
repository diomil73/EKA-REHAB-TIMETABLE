"""Duration-aware afternoon physiotherapy admission preview, with no writes.

Every proposed row must be paired with an explicit, medically approved end
time. Never infer durations from start times, old Excel data or standard slots.
Shared simultaneous groups require a separate exact membership approval.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable, Mapping

from rehab_excel.afternoon_slots_source import (
    AfternoonSlot, AfternoonSlotsError, validate_afternoon_rows,
)
from rehab_excel.afternoon_shared_slot_gate import (
    SharedSlotApproval, SharedSlotGateReport, review_afternoon_shared_slots,
)


class AfternoonIntervalError(ValueError):
    """A clinically material duration or overlap prevents admission."""


@dataclass(frozen=True)
class AfternoonTimedSlot:
    slot: AfternoonSlot
    end_time: time

    def __post_init__(self) -> None:
        if (
            not isinstance(self.slot, AfternoonSlot)
            or not isinstance(self.end_time, time)
            or self.slot.start_time >= self.end_time
        ):
            raise AfternoonIntervalError(
                "Every afternoon session needs its explicit positive same-day end time"
            )


@dataclass(frozen=True)
class AfternoonIntervalReport:
    slot_count: int
    shared_slot_count: int
    overlapping_pairs: int
    ready: bool


def review_afternoon_intervals(
    timed_slots: Iterable[AfternoonTimedSlot],
    *,
    approvals: Iterable[SharedSlotApproval] = (),
) -> AfternoonIntervalReport:
    """Fail closed on intervals, even when start times differ.

    Same-start group treatment is allowed only if every participant shares
    exactly the same end time and the exact member set was approved separately.
    No cross-programme comparisons are performed; AM ΦΘ is independent.
    This does not imply authorization for clinician availability or capacity.
    """
    items = tuple(timed_slots)
    if any(not isinstance(item, AfternoonTimedSlot) for item in items):
        raise AfternoonIntervalError("Fully timed afternoon slots required")
    try:
        validate_afternoon_rows(item.slot for item in items)
    except AfternoonSlotsError as exc:
        raise AfternoonIntervalError("Invalid afternoon slot collection") from exc
    try:
        group_report: SharedSlotGateReport = review_afternoon_shared_slots(
            (item.slot for item in items), approvals
        )
    except ValueError as exc:
        raise AfternoonIntervalError("Shared-slot approval gate failed") from exc
    overlaps = 0
    for i, a in enumerate(items):
        for b in items[i + 1:]:
            if a.slot.weekday != b.slot.weekday:
                continue
            if a.slot.start_time >= b.end_time or b.slot.start_time >= a.end_time:
                continue
            same_patient = a.slot.patient_id == b.slot.patient_id
            same_therapist = (
                a.slot.therapist.strip().casefold()
                == b.slot.therapist.strip().casefold()
            )
            if not same_patient and not same_therapist:
                continue
            # Only an exact simultaneous shared group can cover one interval.
            approved_group_pair = (
                not same_patient
                and same_therapist
                and a.slot.start_time == b.slot.start_time
                and a.end_time == b.end_time
            )
            if not approved_group_pair:
                overlaps += 1
    if overlaps:
        raise AfternoonIntervalError(
            "SAFETY STOP: Overlapping afternoon session intervals"
        )
    return AfternoonIntervalReport(
        slot_count=len(items),
        shared_slot_count=group_report.shared_slot_count,
        overlapping_pairs=0,
        ready=True,
    )


def review_afternoon_timed_edit(
    plan: "AfternoonEditPlan",
    *,
    explicit_end_times: Mapping[AfternoonSlot, time],
    approvals: Iterable[SharedSlotApproval] = (),
) -> AfternoonIntervalReport:
    """Review entire proposed post-edit collection with explicit ends.

    The mapping must cover the full new schedule and contain no stale entries.
    The edit may not drop an old patient's duration implicitly.
    """
    from rehab_excel.afternoon_patient_edit_plan import AfternoonEditPlan

    if not isinstance(plan, AfternoonEditPlan):
        raise AfternoonIntervalError("Validated afternoon edit preview required")
    if not isinstance(explicit_end_times, Mapping):
        raise AfternoonIntervalError("Explicit end-time mapping required")
    if set(explicit_end_times) != set(plan.after):
        raise AfternoonIntervalError("Complete exact end-time mapping required")
    return review_afternoon_intervals(
        (AfternoonTimedSlot(slot, explicit_end_times[slot]) for slot in plan.after),
        approvals=approvals,
    )
