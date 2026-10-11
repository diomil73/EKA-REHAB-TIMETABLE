from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from pathlib import Path
from typing import Iterable

from rehab_core.day_patterns import parse_day_pattern

from .patient_program_draft import build_patient_program_draft
from .patient_recurring_scheduling import (
    RecurringTherapistChoice,
    recurring_therapist_choices,
)
from .reader import read_base_schedule, read_settings
from .workbook_scheduling_candidates import THERAPIST_SCHEDULING_TREATMENTS


DEFAULT_PROVIDER_DAILY_CAPACITY = 6


@dataclass(frozen=True)
class AutomaticTreatmentRequest:
    treatment: str
    day_pattern: str | None = None
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
    return left.start_time == right.start_time and bool(
        left.weekdays.intersection(right.weekdays)
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


def _entry_weekdays(entry) -> frozenset[int]:
    return _request_weekdays(entry.day_pattern)


def _conflicts_existing_patient(option: _Option, base_entries, patient_id: str) -> bool:
    patient_key = _normalize(patient_id)
    treatment_key = _normalize(option.treatment)
    for entry in base_entries:
        if _normalize(entry.patient_id) != patient_key:
            continue
        if _normalize(entry.treatment) == treatment_key:
            continue
        if entry.start_time != option.start_time:
            continue
        if option.weekdays.intersection(_entry_weekdays(entry)):
            return True
    return False


def _provider_load_by_weekday(base_entries, provider: str) -> dict[int, set[time]]:
    provider_key = _normalize(provider)
    loads: dict[int, set[time]] = {day: set() for day in range(5)}
    for entry in base_entries:
        if not entry.therapist_id or _normalize(entry.therapist_id) != provider_key:
            continue
        for weekday in _entry_weekdays(entry):
            loads.setdefault(weekday, set()).add(entry.start_time)
    return loads


def _provider_is_free(
    base_entries,
    provider: str,
    weekdays: frozenset[int],
    slot: time,
) -> tuple[bool, int]:
    loads = _provider_load_by_weekday(base_entries, provider)
    peak = 0
    for weekday in weekdays:
        occupied = loads.get(weekday, set())
        peak = max(peak, len(occupied))
        if slot in occupied:
            return False, peak
        if len(occupied) >= DEFAULT_PROVIDER_DAILY_CAPACITY:
            return False, peak
    return True, peak


def _generic_options(
    workbook_path: str | Path,
    *,
    request_index: int,
    request: AutomaticTreatmentRequest,
    patient_id: str,
    day_patterns: Iterable[str],
    base_entries,
) -> tuple[_Option, ...]:
    draft = build_patient_program_draft(workbook_path, patient_id=patient_id)
    spec = draft.spec_for(request.treatment)
    settings = read_settings(workbook_path)
    slots = (
        settings.reclined_timeslots
        if _normalize(request.treatment) == _normalize("Ανακλινόμενο")
        else settings.standard_timeslots
    )
    providers = spec.provider_choices if spec.requires_provider else ("",)
    options: list[_Option] = []

    for day_pattern in day_patterns:
        weekdays = _request_weekdays(day_pattern)
        for provider in providers:
            for slot in slots:
                workload = 0
                if provider:
                    free, workload = _provider_is_free(
                        base_entries, provider, weekdays, slot
                    )
                    if not free:
                        continue
                option = _Option(
                    request_index=request_index,
                    treatment=request.treatment,
                    day_pattern=day_pattern,
                    weekdays=weekdays,
                    therapist_id=provider,
                    therapist_name=provider,
                    start_time=slot,
                    workload_count=workload,
                )
                if _conflicts_existing_patient(option, base_entries, patient_id):
                    continue
                options.append(option)

    return tuple(options)


def _generic_failure_reason(
    workbook_path: str | Path,
    *,
    request: AutomaticTreatmentRequest,
    patient_id: str,
    day_patterns: Iterable[str],
    base_entries,
) -> str:
    draft = build_patient_program_draft(workbook_path, patient_id=patient_id)
    spec = draft.spec_for(request.treatment)
    settings = read_settings(workbook_path)
    slots = (
        settings.reclined_timeslots
        if _normalize(request.treatment) == _normalize("Ανακλινόμενο")
        else settings.standard_timeslots
    )

    if spec.requires_provider and not spec.provider_choices:
        return (
            "Δεν υπάρχουν καταχωρημένοι providers στο SETTINGS για αυτή τη θεραπεία. "
            "Έλεγξε τη λίστα της ειδικότητας πριν θεωρηθεί ότι δεν υπάρχει διαθέσιμος χώρος."
        )
    if not tuple(day_patterns):
        return "Δεν υπάρχουν day patterns στο SETTINGS για αυτόματο προγραμματισμό."
    if not slots:
        return "Δεν υπάρχουν timeslots στο SETTINGS για αυτή τη θεραπεία."

    if request.provider_id and spec.requires_provider:
        allowed = {_normalize(value) for value in spec.provider_choices}
        if _normalize(request.provider_id) not in allowed:
            return "Ο χειροκίνητα επιλεγμένος provider δεν υπάρχει στη λίστα SETTINGS της θεραπείας."

    providers = spec.provider_choices if spec.requires_provider else ("",)
    any_provider_slot = False
    any_patient_compatible = False
    for day_pattern in day_patterns:
        weekdays = _request_weekdays(day_pattern)
        for provider in providers:
            for slot in slots:
                if request.start_time is not None and slot != request.start_time:
                    continue
                if provider and request.provider_id and _normalize(provider) != _normalize(request.provider_id):
                    continue
                if provider:
                    free, _ = _provider_is_free(base_entries, provider, weekdays, slot)
                    if not free:
                        continue
                any_provider_slot = True
                option = _Option(
                    request_index=-1,
                    treatment=request.treatment,
                    day_pattern=day_pattern,
                    weekdays=weekdays,
                    therapist_id=provider,
                    therapist_name=provider,
                    start_time=slot,
                    workload_count=0,
                )
                if not _conflicts_existing_patient(option, base_entries, patient_id):
                    any_patient_compatible = True
                    break
            if any_patient_compatible:
                break
        if any_patient_compatible:
            break

    if not any_provider_slot and spec.requires_provider:
        return (
            "Υπάρχουν providers στο SETTINGS, αλλά κανένας δεν έχει διαθέσιμο slot "
            "στο ζητούμενο εύρος ημερών/ωρών ή έχει φτάσει το ημερήσιο όριο."
        )
    if any_provider_slot and not any_patient_compatible:
        return "Υπάρχει διαθέσιμος provider/slot, αλλά συγκρούεται με άλλη θεραπεία του ασθενούς."
    return "Δεν βρέθηκε διαθέσιμος συνδυασμός που να τηρεί τις χειροκίνητες επιλογές και τη διαθεσιμότητα."


def _apply_manual_constraints(
    request: AutomaticTreatmentRequest,
    options: Iterable[_Option],
) -> tuple[_Option, ...]:
    constrained = tuple(options)
    if request.day_pattern:
        day_key = _normalize(request.day_pattern)
        constrained = tuple(
            option for option in constrained if _normalize(option.day_pattern) == day_key
        )
    if request.provider_id:
        provider_key = _normalize(request.provider_id)
        constrained = tuple(
            option
            for option in constrained
            if provider_key
            in {
                _normalize(option.therapist_id),
                _normalize(option.therapist_name),
            }
        )
    if request.start_time is not None:
        constrained = tuple(
            option for option in constrained if option.start_time == request.start_time
        )
    return constrained


def _recliner_adjacency_penalty(selected: dict[int, _Option]) -> int:
    physiotherapy = [
        option
        for option in selected.values()
        if _normalize(option.treatment) == _normalize("ΦΘ")
    ]
    recliners = [
        option
        for option in selected.values()
        if _normalize(option.treatment) == _normalize("Ανακλινόμενο")
    ]
    if not physiotherapy or not recliners:
        return 0

    penalty = 0
    for recliner in recliners:
        distances = [
            abs(_minutes(recliner.start_time) - _minutes(physio.start_time))
            for physio in physiotherapy
            if recliner.weekdays.intersection(physio.weekdays)
        ]
        if distances:
            penalty += min(distances)
    return penalty


def propose_automatic_patient_schedule(
    workbook_path: str | Path,
    *,
    patient_id: str,
    requests: Iterable[AutomaticTreatmentRequest],
    anchor_date: date,
    additional_robotic_capable_names: Iterable[str] = (),
) -> AutomaticScheduleProposal:
    """Build a read-only, hybrid recurring patient-program proposal."""

    patient_key = patient_id.strip()
    if not patient_key:
        raise ValueError("patient_id is required")

    request_list = [
        AutomaticTreatmentRequest(
            item.treatment.strip(),
            item.day_pattern.strip() if item.day_pattern else None,
            item.provider_id.strip() if item.provider_id else None,
            item.start_time,
        )
        for item in requests
        if item.treatment.strip()
    ]
    if not request_list:
        raise ValueError("At least one automatic scheduling request is required")

    settings = read_settings(workbook_path)
    base_entries = tuple(read_base_schedule(workbook_path))
    authoritative_treatments = {_normalize(value) for value in settings.therapies}
    physio_supported = {_normalize(value) for value in THERAPIST_SCHEDULING_TREATMENTS}
    fixed_unresolved: dict[int, AutomaticUnresolvedTreatment] = {}
    options_by_request: dict[int, tuple[_Option, ...]] = {}

    for index, request in enumerate(request_list):
        if _normalize(request.treatment) not in authoritative_treatments:
            fixed_unresolved[index] = AutomaticUnresolvedTreatment(
                request.treatment,
                request.day_pattern or "",
                "Η θεραπεία δεν υπάρχει στις Ρυθμίσεις.",
            )
            continue

        candidate_patterns = (
            (request.day_pattern,) if request.day_pattern else tuple(settings.day_patterns)
        )
        options: list[_Option] = []
        generic = _normalize(request.treatment) not in physio_supported

        if not generic:
            for day_pattern in candidate_patterns:
                choices = recurring_therapist_choices(
                    workbook_path,
                    patient_id=patient_key,
                    treatment=request.treatment,
                    day_pattern=day_pattern,
                    anchor_date=anchor_date,
                    additional_robotic_capable_names=additional_robotic_capable_names,
                )
                options.extend(
                    _choice_options(
                        index,
                        request.treatment,
                        day_pattern,
                        choices,
                    )
                )
            options = [
                option
                for option in options
                if not _conflicts_existing_patient(option, base_entries, patient_key)
            ]
        else:
            options.extend(
                _generic_options(
                    workbook_path,
                    request_index=index,
                    request=request,
                    patient_id=patient_key,
                    day_patterns=candidate_patterns,
                    base_entries=base_entries,
                )
            )

        constrained = _apply_manual_constraints(request, options)
        if not constrained:
            reason = (
                _generic_failure_reason(
                    workbook_path,
                    request=request,
                    patient_id=patient_key,
                    day_patterns=candidate_patterns,
                    base_entries=base_entries,
                )
                if generic
                else "Δεν βρέθηκε διαθέσιμος συνδυασμός που να τηρεί τις χειροκίνητες επιλογές και τη διαθεσιμότητα."
            )
            fixed_unresolved[index] = AutomaticUnresolvedTreatment(
                request.treatment,
                request.day_pattern or "",
                reason,
            )
            continue

        options_by_request[index] = tuple(
            sorted(
                constrained,
                key=lambda option: (
                    option.workload_count,
                    _minutes(option.start_time),
                    option.day_pattern.casefold(),
                    option.therapist_name.casefold(),
                ),
            )
        )

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
        adjacency = _recliner_adjacency_penalty(selected)
        total_workload = sum(option.workload_count for option in selected.values())
        time_score = sum(_minutes(option.start_time) for option in selected.values())
        signature = tuple(
            (
                index,
                option.day_pattern.casefold(),
                option.therapist_name.casefold(),
                _minutes(option.start_time),
            )
            for index, option in sorted(selected.items())
        )
        return (len(skipped), adjacency, total_workload, time_score, signature)

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
            request.day_pattern or "",
            "Οι διαθέσιμες επιλογές συγκρούονται χρονικά με άλλη θεραπεία του ασθενούς.",
        )

    unresolved = tuple(
        unresolved_by_index[index] for index in sorted(unresolved_by_index)
    )
    return AutomaticScheduleProposal(
        patient_id=patient_key,
        anchor_date=anchor_date,
        assignments=assignments,
        unresolved=unresolved,
        total_workload=sum(item.workload_count for item in assignments),
    )
