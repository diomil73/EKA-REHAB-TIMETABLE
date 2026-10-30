from __future__ import annotations

from dataclasses import replace
from datetime import date
from typing import Iterable
import unicodedata

from .base_schedule import pattern_applies_on_date
from .models import BaseScheduleEntry, Session
from .replacement_options import ReplacementProviderOption
from .service_hours import POOL_SERVICE_ID, ROBOTIC_SERVICE_ID, SpecialServiceHoursBook


def _norm(value: str | None) -> str:
    decomposed = unicodedata.normalize("NFD", str(value or "").strip().casefold())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def build_service_id_by_session(
    base_entries: Iterable[BaseScheduleEntry],
    target_date: date,
) -> dict[str, str]:
    """Build explicit special-service identities for materialized sessions.

    Robotic identity already survives materialization on ``Session.robotic`` and
    therefore needs no mapping. Pool identity does not yet have a dedicated
    session field, so it is carried explicitly from the authoritative recurring
    entry instead of being guessed later from display text.
    """

    mapping: dict[str, str] = {}
    for entry in base_entries:
        if entry.therapist_id is None:
            continue
        if not pattern_applies_on_date(entry.day_pattern, target_date):
            continue
        if _norm(entry.treatment) != _norm("Πισίνα"):
            continue
        mapping[f"{entry.base_entry_id}@{target_date.isoformat()}"] = POOL_SERVICE_ID
    return mapping


def service_id_for_session(
    session: Session,
    *,
    explicit_service_id: str | None = None,
) -> str | None:
    """Resolve the special service whose operating hours constrain a session.

    Robotic identity is authoritative on ``Session.robotic``. Other special
    services, such as pool/hydrotherapy, must be supplied explicitly by the
    caller from their structured authoritative source. Do not infer them from
    display text at this stage.
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
