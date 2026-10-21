from __future__ import annotations

from dataclasses import replace
from typing import Iterable

from .models import Session
from .replacement_options import ReplacementProviderOption
from .service_hours import ROBOTIC_SERVICE_ID, SpecialServiceHoursBook


def service_id_for_session(
    session: Session,
    *,
    explicit_service_id: str | None = None,
) -> str | None:
    """Resolve the special service whose operating hours constrain a session.

    Robotic identity is authoritative on ``Session.robotic``. Other special
    services, such as pool/hydrotherapy, must be supplied explicitly by the
    caller until their own structured session identity exists. Do not infer
    them from display text.
    """

    if explicit_service_id:
        return explicit_service_id
    if session.robotic:
        return ROBOTIC_SERVICE_ID
    return None


def filter_replacement_options_for_service_hours(
    options: Iterable[ReplacementProviderOption],
    *,
    target_session: Session,
    hours_book: SpecialServiceHoursBook,
    explicit_service_id: str | None = None,
) -> list[ReplacementProviderOption]:
    service_id = service_id_for_session(
        target_session,
        explicit_service_id=explicit_service_id,
    )
    if service_id is None:
        return list(options)

    allowed = set(hours_book.allowed_times(service_id, target_session.session_date))
    filtered: list[ReplacementProviderOption] = []
    for option in options:
        times = tuple(slot for slot in option.available_timeslots if slot in allowed)
        if not times:
            continue
        filtered.append(
            replace(
                option,
                recommended_time=times[0],
                exact_time_available=option.requested_time in times,
                available_timeslots=times,
            )
        )
    return filtered
