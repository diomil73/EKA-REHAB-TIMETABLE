from __future__ import annotations

from .models import DailySessionCancellation, Session, SessionCancellationKind


TREATMENT_SUSPENSION_REASON = "Αναστολή θεραπείας"


def create_treatment_suspension(
    session: Session,
    *,
    cancellation_id: str | None = None,
) -> DailySessionCancellation:
    """Resolve one concrete orphan by suspending only that day's treatment.

    The recurring programme is not modified. The returned cancellation is a
    date-specific operational fact that causes the daily state to become
    CANCELLED, releases the original provider slot, and remains auditable for
    statistics.
    """

    return DailySessionCancellation(
        cancellation_id=(
            cancellation_id
            or f"suspend:{session.session_id}:{session.session_date.isoformat()}"
        ),
        target_session_id=session.session_id,
        cancellation_date=session.session_date,
        kind=SessionCancellationKind.TREATMENT_SUSPENDED,
        reason=TREATMENT_SUSPENSION_REASON,
    )
