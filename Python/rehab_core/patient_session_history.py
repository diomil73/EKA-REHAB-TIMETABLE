from __future__ import annotations

from dataclasses import dataclass
from datetime import date, time
from typing import Iterable

from .daily_state import DailySessionState, DailySessionStatus


class PatientSessionHistoryError(RuntimeError):
    """Raised when immutable patient/session history cannot be frozen safely."""


@dataclass(frozen=True)
class PatientSessionHistoryRow:
    """Immutable outcome of one patient session inside one finalized snapshot."""

    snapshot_id: str
    service_date: date
    snapshot_version: int
    session_id: str
    patient_id: str
    status: DailySessionStatus
    original_therapist_id: str
    original_time: time
    effective_therapist_id: str | None
    effective_time: time | None
    replacement_id: str | None
    cancellation_id: str | None
    cancellation_kind: str | None
    reason: str | None


def build_patient_session_history_rows(
    states: Iterable[DailySessionState],
    *,
    snapshot_id: str,
    service_date: date,
    snapshot_version: int,
) -> tuple[PatientSessionHistoryRow, ...]:
    """Freeze patient/session outcomes from one final effective daily state.

    The rows deliberately retain both original and effective provider/time so
    later reports can distinguish scheduled treatment from delivered treatment
    without consulting mutable operational sheets.
    """

    snapshot_id = snapshot_id.strip()
    if not snapshot_id:
        raise PatientSessionHistoryError("snapshot_id cannot be blank")
    if snapshot_version < 1:
        raise PatientSessionHistoryError("snapshot_version must be at least 1")

    state_list = tuple(states)
    wrong_date = tuple(
        state.session_id for state in state_list if state.session_date != service_date
    )
    if wrong_date:
        raise PatientSessionHistoryError(
            "Patient/session history received states from another service date: "
            + ", ".join(wrong_date)
        )

    unresolved = tuple(
        state.session_id
        for state in state_list
        if state.status == DailySessionStatus.THERAPIST_ABSENT
    )
    if unresolved:
        raise PatientSessionHistoryError(
            "Cannot freeze patient/session history with unresolved sessions: "
            + ", ".join(unresolved)
        )

    seen_session_ids: set[str] = set()
    rows: list[PatientSessionHistoryRow] = []
    for state in sorted(
        state_list,
        key=lambda item: (
            item.original_time,
            item.patient_id.casefold(),
            item.session_id.casefold(),
        ),
    ):
        session_id = state.session_id.strip()
        patient_id = state.patient_id.strip()
        if not session_id:
            raise PatientSessionHistoryError("session_id cannot be blank")
        if not patient_id:
            raise PatientSessionHistoryError(
                f"Patient id is blank for session {state.session_id!r}"
            )
        session_key = session_id.casefold()
        if session_key in seen_session_ids:
            raise PatientSessionHistoryError(
                f"Duplicate session_id in finalized state: {state.session_id!r}"
            )
        seen_session_ids.add(session_key)

        rows.append(
            PatientSessionHistoryRow(
                snapshot_id=snapshot_id,
                service_date=service_date,
                snapshot_version=snapshot_version,
                session_id=session_id,
                patient_id=patient_id,
                status=state.status,
                original_therapist_id=state.original_therapist_id,
                original_time=state.original_time,
                effective_therapist_id=state.effective_therapist_id,
                effective_time=state.effective_time,
                replacement_id=state.replacement_id,
                cancellation_id=state.cancellation_id,
                cancellation_kind=(
                    state.cancellation_kind.value
                    if state.cancellation_kind is not None
                    else None
                ),
                reason=state.reason,
            )
        )

    return tuple(rows)
