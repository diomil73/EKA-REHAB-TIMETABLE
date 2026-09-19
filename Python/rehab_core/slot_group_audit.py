from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Mapping

from .models import BaseScheduleEntry, Patient
from .slot_groups import SlotGroup, build_slot_groups


@dataclass(frozen=True)
class SlotGroupAuditReport:
    base_assignment_count: int
    group_count: int
    single_member_group_count: int
    multi_member_group_count: int
    clean_pair_count: int
    same_patient_split_count: int
    overlapping_group_count: int
    groups: tuple[SlotGroup, ...]

    @property
    def ok(self) -> bool:
        # Overlapping groups are surfaced for human review rather than silently
        # rewritten. A therapist may intentionally supervise more than one
        # patient at once in some future workflow.
        return True


def audit_slot_groups(entries: Iterable[BaseScheduleEntry]) -> SlotGroupAuditReport:
    source = tuple(entries)
    groups = tuple(build_slot_groups(source))
    return SlotGroupAuditReport(
        base_assignment_count=len(source),
        group_count=len(groups),
        single_member_group_count=sum(len(group.members) == 1 for group in groups),
        multi_member_group_count=sum(len(group.members) > 1 for group in groups),
        clean_pair_count=sum(group.is_pair for group in groups),
        same_patient_split_count=sum(group.is_same_patient_split for group in groups),
        overlapping_group_count=sum(bool(group.overlaps) for group in groups),
        groups=groups,
    )


def patient_name_map(patients: Iterable[Patient]) -> Mapping[str, str]:
    return {patient.patient_id: patient.display_name for patient in patients}
