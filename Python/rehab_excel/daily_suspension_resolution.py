from __future__ import annotations

from datetime import date
from typing import Iterable

from rehab_core.models import (
    DailySessionCancellation,
    Session,
    SessionCancellationKind,
)


class DailySuspensionResolutionError(RuntimeError):
    """Raised when unresolved sessions cannot be converted safely to suspensions."""


SUSPENSION_STATUS = "Αναστολή θεραπείας"
SUSPENSION_REASON = "Δεν βρέθηκε έγκυρη αναπλήρωση από τον global solver"


def build_unresolved_suspensions(
    unresolved_request_ids: Iterable[str],
    *,
    sessions: Iterable[Session],
    target_date: date,
) -> tuple[DailySessionCancellation, ...]:
    """Convert solver leftovers into explicit date-specific treatment suspensions."""

    session_by_id = {session.session_id: session for session in sessions}
    requested = tuple(
        dict.fromkeys(
            str(value).strip()
            for value in unresolved_request_ids
            if str(value).strip()
        )
    )
    suspensions: list[DailySessionCancellation] = []

    for session_id in requested:
        session = session_by_id.get(session_id)
        if session is None:
            raise DailySuspensionResolutionError(
                f"Unresolved session {session_id!r} does not exist in the materialized day"
            )
        if session.session_date != target_date:
            raise DailySuspensionResolutionError(
                f"Unresolved session {session_id!r} belongs to {session.session_date.isoformat()}, "
                f"not {target_date.isoformat()}"
            )
        suspensions.append(
            DailySessionCancellation(
                cancellation_id=f"global-suspension:{session_id}:{target_date.isoformat()}",
                target_session_id=session_id,
                cancellation_date=target_date,
                kind=SessionCancellationKind.TREATMENT_SUSPENDED,
                reason=f"{SUSPENSION_STATUS} | {SUSPENSION_REASON}",
            )
        )

    return tuple(suspensions)
