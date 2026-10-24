from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from enum import Enum
from typing import Iterable

from .models import (
    AbsenceKind,
    DailyAbsence,
    DailySessionCancellation,
    ReplacementAssignment,
    ReplacementProviderKind,
    Session,
    SessionCancellationKind,
)


class DailySessionStatus(str, Enum):
    ACTIVE = "active"
    PATIENT_ABSENT = "patient_absent"
    THERAPIST_ABSENT = "therapist_absent"
    REPLACED = "replaced"
    CANCELLED = "cancelled"


@dataclass(frozen=True)
class DailySessionState:
    """Effective presentation state for one immutable base session.

    The original therapist/time are always preserved. Effective therapist/time
    describe what should be shown operationally for the day. This is the bridge
    between scheduling logic and the future Excel rendering layer.
    """

    session_id: str
    patient_id: str
    session_date: date
    status: DailySessionStatus
    original_therapist_id: str
    original_time: time
    effective_therapist_id: str | None
    effective_time: time | None
    replacement_id: str | None = None
    reason: str | None = None
    effective_provider_kind: ReplacementProviderKind | None = None
    cancellation_id: str | None = None
    cancellation_kind: SessionCancellationKind | None = None

    @property
    def needs_replacement(self) -> bool:
        return self.status == DailySessionStatus.THERAPIST_ABSENT

    @property
    def releases_provider_slot(self) -> bool:
        """Whether the original therapist/time may be reused for another patient."""

        return self.status in {
            DailySessionStatus.PATIENT_ABSENT,
            DailySessionStatus.CANCELLED,
        }

    @property
    def original_should_be_struck_through(self) -> bool:
        return self.status in {
            DailySessionStatus.PATIENT_ABSENT,
            DailySessionStatus.THERAPIST_ABSENT,
            DailySessionStatus.CANCELLED,
        }


def _covering_absence(
    *,
    kind: AbsenceKind,
    subject_id: str,
    target_date: date,
    target_time: time,
    absences: Iterable[DailyAbsence],
) -> DailyAbsence | None:
    for absence in absences:
        if (
            absence.absence_kind == kind
            and absence.subject_id == subject_id
            and absence.covers(target_date, target_time)
        ):
            return absence
    return None


def build_daily_session_states(
    sessions: Iterable[Session],
    absences: Iterable[DailyAbsence] = (),
    replacements: Iterable[ReplacementAssignment] = (),
    cancellations: Iterable[DailySessionCancellation] = (),
    *,
    target_date: date | None = None,
) -> list[DailySessionState]:
    """Build an operational view while keeping every base Session untouched.

    Priority for a base session is:
    1. explicit session cancellation/no-show;
    2. patient absence, because no treatment can take place;
    3. accepted replacement overlay;
    4. therapist absence, which creates a pending replacement need;
    5. normal active session.

    A cancellation removes only that day's concrete treatment from effective
    workload. The recurring/base session is not edited and the original
    therapist slot becomes available for other work or a replacement patient.

    ``target_date`` optionally limits the result to one working day.
    """

    sessions = tuple(sessions)
    absences = tuple(absences)
    replacements = tuple(replacements)
    cancellations = tuple(cancellations)
    replacement_by_session = {
        replacement.target_session_id: replacement for replacement in replacements
    }
    cancellation_by_session = {
        cancellation.target_session_id: cancellation for cancellation in cancellations
    }

    states: list[DailySessionState] = []

    for session in sessions:
        if target_date is not None and session.session_date != target_date:
            continue

        cancellation = cancellation_by_session.get(session.session_id)
        if cancellation is not None and cancellation.applies_to(session):
            states.append(
                DailySessionState(
                    session_id=session.session_id,
                    patient_id=session.patient_id,
                    session_date=session.session_date,
                    status=DailySessionStatus.CANCELLED,
                    original_therapist_id=session.therapist_id,
                    original_time=session.start_time,
                    effective_therapist_id=None,
                    effective_time=None,
                    reason=cancellation.reason,
                    cancellation_id=cancellation.cancellation_id,
                    cancellation_kind=cancellation.kind,
                )
            )
            continue

        patient_absence = _covering_absence(
            kind=AbsenceKind.PATIENT,
            subject_id=session.patient_id,
            target_date=session.session_date,
            target_time=session.start_time,
            absences=absences,
        )
        if patient_absence is not None:
            states.append(
                DailySessionState(
                    session_id=session.session_id,
                    patient_id=session.patient_id,
                    session_date=session.session_date,
                    status=DailySessionStatus.PATIENT_ABSENT,
                    original_therapist_id=session.therapist_id,
                    original_time=session.start_time,
                    effective_therapist_id=None,
                    effective_time=None,
                    reason=patient_absence.reason,
                )
            )
            continue

        replacement = replacement_by_session.get(session.session_id)
        if replacement is not None:
            replacement_patient_absence = _covering_absence(
                kind=AbsenceKind.PATIENT,
                subject_id=session.patient_id,
                target_date=replacement.replacement_date,
                target_time=replacement.replacement_time,
                absences=absences,
            )
            if replacement_patient_absence is not None:
                states.append(
                    DailySessionState(
                        session_id=session.session_id,
                        patient_id=session.patient_id,
                        session_date=session.session_date,
                        status=DailySessionStatus.PATIENT_ABSENT,
                        original_therapist_id=session.therapist_id,
                        original_time=session.start_time,
                        effective_therapist_id=None,
                        effective_time=None,
                        replacement_id=replacement.replacement_id,
                        reason=replacement_patient_absence.reason,
                    )
                )
                continue

            states.append(
                DailySessionState(
                    session_id=session.session_id,
                    patient_id=session.patient_id,
                    session_date=session.session_date,
                    status=DailySessionStatus.REPLACED,
                    original_therapist_id=session.therapist_id,
                    original_time=session.start_time,
                    effective_therapist_id=replacement.replacement_therapist_id,
                    effective_time=replacement.replacement_time,
                    replacement_id=replacement.replacement_id,
                    reason=replacement.reason,
                    effective_provider_kind=replacement.replacement_provider_kind,
                )
            )
            continue

        therapist_absence = _covering_absence(
            kind=AbsenceKind.THERAPIST,
            subject_id=session.therapist_id,
            target_date=session.session_date,
            target_time=session.start_time,
            absences=absences,
        )
        if therapist_absence is not None:
            states.append(
                DailySessionState(
                    session_id=session.session_id,
                    patient_id=session.patient_id,
                    session_date=session.session_date,
                    status=DailySessionStatus.THERAPIST_ABSENT,
                    original_therapist_id=session.therapist_id,
                    original_time=session.start_time,
                    effective_therapist_id=None,
                    effective_time=None,
                    reason=therapist_absence.reason,
                )
            )
            continue

        states.append(
            DailySessionState(
                session_id=session.session_id,
                patient_id=session.patient_id,
                session_date=session.session_date,
                status=DailySessionStatus.ACTIVE,
                original_therapist_id=session.therapist_id,
                original_time=session.start_time,
                effective_therapist_id=session.therapist_id,
                effective_time=session.start_time,
                effective_provider_kind=ReplacementProviderKind.THERAPIST,
            )
        )

    states.sort(key=lambda state: (state.session_date, state.original_time, state.session_id))
    return states
