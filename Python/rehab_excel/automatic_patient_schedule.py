from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from pathlib import Path
from typing import Iterable

from rehab_core.day_patterns import parse_day_pattern

from .patient_recurring_scheduling import (
    RecurringTherapistChoice,
    recurring_therapist_choices,
)
from .workbook_scheduling_candidates import THERAPIST_SCHEDULING_TREATMENTS


@dataclass(frozen=True)
class AutomaticTreatmentRequest:
    treatment: str
    day_pattern: str
    provider_id: str | None = None
    start_time: time | None = None


@dataclass(frozen=True)
class AutomaticProposedAssignment:
    treatment: str
    day_pattern: str
    therapist_id: str
    therapist_name: str
    start_time: time
    workload_count: int


@dataclass(frozen=True)
class AutomaticUnresolvedTreatment:
    treatment: str
    day_pattern: str
    reason: str


@dataclass(frozen=True)
class AutomaticScheduleProposal:
    patient_id: str
    anchor_date: date
    assignments: tuple[AutomaticProposedAssignment, ...]
    unresolved: tuple[AutomaticUnresolvedTreatment, ...]
    total_workload: int

    @property
    def complete(self) -> bool:
        return not self.unresolved


@dataclass(frozen=True)
class _Option:
    request_index: int
    treatment: str
    day_pattern: str
    weekdays: frozenset[int]
    therapist_id: str
    therapist_name: str
    start_time: time
    workload_count: int


def _normalize(value: str) -> str:
    return value.strip().casefold()


def _minutes(value: time) -> int:
    return value.hour * 60 + value.minute


def _request_weekdays(day_pattern: str) -> frozenset[int]:
    return frozenset(int(day) for day in parse_day_pattern(day_pattern))


def _patient_conflict(left: _Option, right: _Option) -> bool:
    return (
        left.start_time == right.start_time
        and bool(left.weekdays.intersection(right.weekdays))
    )


def _choice_options(
    request_index: int,
    treatment: str,
    day_pattern: str,
    choices: Iterable[RecurringTherapistChoice],
) -> tuple[_Option, ...]:
    weekdays = _request_weekdays(day_pattern)
    options = [
        _Option(
            request_index=request_index,
            treatment=treatment,
            day_pattern=day_pattern,
            weekdays=weekdays,
            therapist_id=choice.therapist_id,
            therapist_name=choice.display_name,
            start_time=slot,
            workload_count=choice.workload_count,
        )
        for choice in choices
        for slot in choice.available_slots
    ]
    options.sort(
        key=lambda option: (
            option.workload_count,
            _minutes(option.start_time),
            option.therapist_name.casefold(),
            option.therapist_id,
        )
    )
    return tuple(options)


def _apply_manual_constraints(
    request: AutomaticTreatmentRequest,
    options: Iterable[_Option],
) -> tuple[_Option, ...]:
    constrained = tuple(options)
    if request.provider_id:
        provider_key = _normalize(request.provider_id)
        constrained = tuple(
            option
            for option in constrained
            if provider_key in {
                _normalize(option.therapist_id),
                _normalize(option.therapist_name),
            }
        )
    if request.start_time is not None:
        constrained = tuple(
            option for option in constrained if option.start_time == request.start_time
        )
    return constrained


def propose_automatic_patient_schedule(
    workbook_path: str | Path,
    *,
    patient_id: str,
    requests: Iterable[AutomaticTreatmentRequest],
    anchor_date: date,
    additional_robotic_capable_names: Iterable[str] = (),
) -> AutomaticScheduleProposal:
    """Build a read-only best-effort recurring schedule proposal.

    The function never writes to the workbook. It reuses the authoritative
    recurring candidate engine for every supported treatment, then searches for
    a combination that maximizes the number of scheduled requests without
    placing the patient in two overlapping treatments at the same time.

    Manual provider/time values on a request are treated as hard constraints:
    the solver may fill missing fields but never silently replace user choices.

    At this stage the common therapist engine supports ΦΘ and Ρομποτικό only.
    Other treatment families are returned explicitly as unresolved rather than
    being guessed through physiotherapist rules that do not apply to them.
    """

    patient_key = patient_id.strip()
    if not patient_key:
        raise ValueError("patient_id is required")

    request_list = [
        AutomaticTreatmentRequest(
            item.treatment.strip(),
            item.day_pattern.strip(),
            item.provider_id.strip() if item.provider_id else None,
            item.start_time,
        )
        for item in requests
        if item.treatment.strip() or item.day_pattern.strip()
    ]
    if not request_list:
        raise ValueError("At least one automatic scheduling request is required")
    if any(not item.treatment or not item.day_pattern for item in request_list):
        raise ValueError("Each automatic scheduling request requires treatment and day_pattern")

    supported = {_normalize(value) for value in THERAPIST_SCHEDULING_TREATMENTS}
    fixed_unresolved: dict[int, AutomaticUnresolvedTreatment] = {}
    options_by_request: dict[int, tuple[_Option, ...]] = {}

    for index, request in enumerate(request_list):
        if _normalize(request.treatment) not in supported:
            fixed_unresolved[index] = AutomaticUnresolvedTreatment(
                request.treatment,
                request.day_pattern,
                "Δεν υποστηρίζεται ακόμη από τον automatic physiotherapist solver.",
            )
            continue

        choices = recurring_therapist_choices(
            workbook_path,
            patient_id=patient_key,
            treatment=request.treatment,
            day_pattern=request.day_pattern,
            anchor_date=anchor_date,
            additional_robotic_capable_names=additional_robotic_capable_names,
        )
        options = _choice_options(
            index,
            request.treatment,
            request.day_pattern,
            choices,
        )
        if not options:
            fixed_unresolved[index] = AutomaticUnresolvedTreatment(
                request.treatment,
                request.day_pattern,
                "Δεν βρέθηκε κοινός διαθέσιμος θεραπευτής/ώρα για όλο το μοτίβο ημερών.",
            )
            continue

        constrained_options = _apply_manual_constraints(request, options)
        if not constrained_options:
            fixed_unresolved[index] = AutomaticUnresolvedTreatment(
                request.treatment,
                request.day_pattern,
                "Οι χειροκίνητες επιλογές θεραπευτή/ώρας δεν είναι διαθέσιμες για όλο το μοτίβο ημερών.",
            )
            continue
        options_by_request[index] = constrained_options

    searchable_indices = sorted(
        options_by_request,
        key=lambda index: (
            len(options_by_request[index]),
            0 if _normalize(request_list[index].treatment) == _normalize("Ρομποτικό") else 1,
            index,
        ),
    )

    best_selected: dict[int, _Option] | None = None
    best_skipped: tuple[int, ...] | None = None
    best_score: tuple | None = None

    def score(selected: dict[int, _Option], skipped: tuple[int, ...]) -> tuple:
        total_workload = sum(option.workload_count for option in selected.values())
        time_score = sum(_minutes(option.start_time) for option in selected.values())
        signature = tuple(
            (
                index,
                option.therapist_name.casefold(),
                option.therapist_id,
                _minutes(option.start_time),
            )
            for index, option in sorted(selected.items())
        )
        return (len(skipped), total_workload, time_score, signature)

    def search(position: int, selected: dict[int, _Option], skipped: tuple[int, ...]) -> None:
        nonlocal best_selected, best_skipped, best_score
        if position >= len(searchable_indices):
            current_score = score(selected, skipped)
            if best_score is None or current_score < best_score:
                best_score = current_score
                best_selected = dict(selected)
                best_skipped = skipped
            return

        index = searchable_indices[position]
        for option in options_by_request[index]:
            if any(_patient_conflict(option, existing) for existing in selected.values()):
                continue
            selected[index] = option
            search(position + 1, selected, skipped)
            selected.pop(index, None)

        # Best-effort mode: a partial proposal is better than no proposal when
        # all combinations conflict. The score always penalizes skipping before
        # workload/time, so a complete feasible proposal wins whenever one exists.
        search(position + 1, selected, (*skipped, index))

    search(0, {}, ())
    selected = best_selected or {}
    skipped = set(best_skipped or ())

    assignments = tuple(
        AutomaticProposedAssignment(
            treatment=option.treatment,
            day_pattern=option.day_pattern,
            therapist_id=option.therapist_id,
            therapist_name=option.therapist_name,
            start_time=option.start_time,
            workload_count=option.workload_count,
        )
        for _, option in sorted(selected.items())
    )

    unresolved_by_index = dict(fixed_unresolved)
    for index in skipped:
        request = request_list[index]
        unresolved_by_index[index] = AutomaticUnresolvedTreatment(
            request.treatment,
            request.day_pattern,
            "Οι διαθέσιμες επιλογές συγκρούονται χρονικά με άλλη προτεινόμενη θεραπεία του ασθενούς.",
        )

    unresolved = tuple(
        unresolved_by_index[index]
        for index in sorted(unresolved_by_index)
    )
    return AutomaticScheduleProposal(
        patient_id=patient_key,
        anchor_date=anchor_date,
        assignments=assignments,
        unresolved=unresolved,
        total_workload=sum(item.workload_count for item in assignments),
    )
