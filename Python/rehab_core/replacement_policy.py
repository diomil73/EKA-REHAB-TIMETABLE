from __future__ import annotations

from dataclasses import replace
from datetime import date, time
from typing import Iterable

from .models import (
    BaseScheduleEntry,
    DailyAbsence,
    Patient,
    ReplacementAssignment,
    Session,
    Student,
    StudentAssignment,
    Therapist,
)
from .patient_schedule import filter_replacement_options_for_patient_schedule
from .provider_policy import ProviderPolicyBook, ProviderProfile, ProviderType
from .replacement_options import (
    ReplacementProviderOption,
    find_replacement_options,
    order_timeslots_by_preference,
)


def policy_adjusted_therapists(
    therapists: Iterable[Therapist],
    *,
    policy_book: ProviderPolicyBook,
    target_date: date,
) -> tuple[Therapist, ...]:
    """Return therapist models adjusted for the effective policy on one date."""

    adjusted: list[Therapist] = []
    for therapist in therapists:
        if therapist.therapist_id not in policy_book.profiles:
            policy_book.profiles[therapist.therapist_id] = ProviderProfile(
                provider_id=therapist.therapist_id,
                display_name=therapist.display_name,
                provider_type=ProviderType.THERAPIST,
                default_max_slots=therapist.max_daily_timeslots,
            )

        state = policy_book.effective_state(therapist.therapist_id, target_date)
        if not state.active:
            continue

        if state.is_active_leader:
            if not state.explicitly_open_replacement_times:
                continue
            effective_limit = len(state.explicitly_open_replacement_times)
        else:
            if not state.replacement_eligible:
                continue
            effective_limit = state.max_slots

        adjusted.append(
            replace(therapist, max_daily_timeslots=max(effective_limit, 0))
        )

    return tuple(adjusted)


def _filter_option_times(
    option: ReplacementProviderOption,
    *,
    policy_book: ProviderPolicyBook,
    target_date: date,
) -> ReplacementProviderOption | None:
    state = policy_book.effective_state(option.provider_id, target_date)
    if not state.active:
        return None

    allowed = tuple(
        slot for slot in option.available_timeslots if slot not in state.blocked_times
    )

    if state.is_active_leader:
        allowed = tuple(
            slot
            for slot in allowed
            if slot in state.explicitly_open_replacement_times
        )
    elif not state.replacement_eligible:
        return None

    if not allowed:
        return None

    ordered = order_timeslots_by_preference(allowed, option.requested_time)
    return replace(
        option,
        recommended_time=ordered[0],
        exact_time_available=option.requested_time in ordered,
        available_timeslots=ordered,
    )


def find_policy_replacement_options(
    target_session: Session,
    therapists: Iterable[Therapist],
    sessions: Iterable[Session],
    *,
    policy_book: ProviderPolicyBook,
    absences: Iterable[DailyAbsence] = (),
    patients: Iterable[Patient] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    requested_time: time | None = None,
    timeslots: Iterable[time] = (),
    students: Iterable[Student] = (),
    student_assignments: Iterable[StudentAssignment] = (),
    base_entries: Iterable[BaseScheduleEntry] = (),
) -> list[ReplacementProviderOption]:
    """Run replacement ranking with provider and patient-schedule policy applied.

    Ranking remains unchanged: daily load first, exact requested time second,
    then the existing tie-breakers. Provider policy and the patient's other
    specialties act only as eligibility/timeslot gates around that ranking.
    """

    adjusted = policy_adjusted_therapists(
        therapists,
        policy_book=policy_book,
        target_date=target_session.session_date,
    )
    options = find_replacement_options(
        target_session=target_session,
        therapists=adjusted,
        sessions=sessions,
        absences=absences,
        patients=patients,
        replacements=replacements,
        requested_time=requested_time,
        timeslots=timeslots,
        students=students,
        student_assignments=student_assignments,
    )

    policy_filtered: list[ReplacementProviderOption] = []
    for option in options:
        revised = _filter_option_times(
            option,
            policy_book=policy_book,
            target_date=target_session.session_date,
        )
        if revised is not None:
            policy_filtered.append(revised)

    return filter_replacement_options_for_patient_schedule(
        policy_filtered,
        target_session=target_session,
        base_entries=base_entries,
    )


def validate_policy_replacement_time(
    *,
    provider_id: str,
    target_date: date,
    replacement_time: time,
    policy_book: ProviderPolicyBook,
) -> None:
    """Safety gate for a selected replacement before creating the overlay."""

    state = policy_book.effective_state(provider_id, target_date)
    if not state.active:
        raise ValueError("Replacement provider is inactive by provider policy")
    if replacement_time in state.blocked_times:
        raise ValueError("Replacement time is blocked by provider policy")
    if state.is_active_leader:
        if replacement_time not in state.explicitly_open_replacement_times:
            raise ValueError(
                "Active manager/acting manager time is not explicitly open for replacement"
            )
    elif not state.replacement_eligible:
        raise ValueError("Replacement provider is excluded by provider policy")
