from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable

from .daily_state import DailySessionStatus, build_daily_session_states
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
from .provider_policy import ProviderPolicyBook
from .replacement_options import ReplacementProviderOption, find_replacement_options
from .replacement_policy import find_policy_replacement_options


@dataclass(frozen=True)
class TherapistAbsenceReplacementItem:
    """One patient session that needs a replacement because its therapist is absent."""

    session_id: str
    patient_id: str
    patient_name: str
    original_therapist_id: str
    original_time: time
    treatment: str | None
    robotic: bool
    options: tuple[ReplacementProviderOption, ...]


@dataclass(frozen=True)
class TherapistAbsenceReplacementQueue:
    items: tuple[TherapistAbsenceReplacementItem, ...]
    skipped_patient_absent: int = 0

    @property
    def affected_sessions(self) -> int:
        return len(self.items)


def build_therapist_absence_replacement_queue(
    *,
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence],
    therapists: Iterable[Therapist],
    patients: Iterable[Patient],
    timeslots: Iterable[time],
    replacements: Iterable[ReplacementAssignment] = (),
    students: Iterable[Student] = (),
    student_assignments: Iterable[StudentAssignment] = (),
    policy_book: ProviderPolicyBook | None = None,
    base_entries: Iterable[BaseScheduleEntry] = (),
) -> TherapistAbsenceReplacementQueue:
    """Build replacement suggestions for sessions affected by therapist absence.

    Patient absence has priority. Provider policy is optional. When recurring
    base entries are supplied, times occupied by any other specialty for the
    same patient are removed from replacement options as well.
    """

    session_list = tuple(sessions)
    absence_list = tuple(absences)
    patient_list = tuple(patients)
    therapist_list = tuple(therapists)
    replacement_list = tuple(replacements)
    student_list = tuple(students)
    student_assignment_list = tuple(student_assignments)
    timeslot_list = tuple(timeslots)
    base_entry_list = tuple(base_entries)

    patient_by_id = {patient.patient_id: patient for patient in patient_list}
    session_by_id = {session.session_id: session for session in session_list}

    states = build_daily_session_states(
        session_list,
        absences=absence_list,
        replacements=replacement_list,
    )
    skipped_patient_absent = sum(
        1 for state in states if state.status == DailySessionStatus.PATIENT_ABSENT
    )

    items: list[TherapistAbsenceReplacementItem] = []
    for state in states:
        if state.status != DailySessionStatus.THERAPIST_ABSENT:
            continue
        session = session_by_id[state.session_id]
        patient = patient_by_id.get(session.patient_id)
        patient_name = patient.display_name if patient is not None else session.patient_id

        if policy_book is None:
            options = find_replacement_options(
                target_session=session,
                therapists=therapist_list,
                sessions=session_list,
                absences=absence_list,
                patients=patient_list,
                replacements=replacement_list,
                requested_time=session.start_time,
                timeslots=timeslot_list,
                students=student_list,
                student_assignments=student_assignment_list,
            )
            options = filter_replacement_options_for_patient_schedule(
                options,
                target_session=session,
                base_entries=base_entry_list,
            )
        else:
            options = find_policy_replacement_options(
                target_session=session,
                therapists=therapist_list,
                sessions=session_list,
                policy_book=policy_book,
                absences=absence_list,
                patients=patient_list,
                replacements=replacement_list,
                requested_time=session.start_time,
                timeslots=timeslot_list,
                students=student_list,
                student_assignments=student_assignment_list,
                base_entries=base_entry_list,
            )
        items.append(
            TherapistAbsenceReplacementItem(
                session_id=session.session_id,
                patient_id=session.patient_id,
                patient_name=patient_name,
                original_therapist_id=session.therapist_id,
                original_time=session.start_time,
                treatment=session.treatment,
                robotic=session.robotic,
                options=tuple(options),
            )
        )

    items.sort(
        key=lambda item: (
            item.original_therapist_id.casefold(),
            item.original_time,
            item.patient_name.casefold(),
            item.session_id,
        )
    )
    return TherapistAbsenceReplacementQueue(
        items=tuple(items),
        skipped_patient_absent=skipped_patient_absent,
    )
