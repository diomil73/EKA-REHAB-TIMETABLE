from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Mapping

from .outpatient_schedule_source import read_unified_base_schedule
from .outpatient_schedule_suggestions import (
    SUGGESTION_MODE_LABELS,
    SuggestionMode,
    suggest_outpatient_schedule_slots,
)
from .reader import read_settings


class OutpatientScheduleSuggestionBridgeError(RuntimeError):
    """Stable UI-facing error for outpatient schedule suggestion requests."""


def _required_text(values: Mapping[str, object], key: str) -> str:
    value = str(values.get(key, "")).strip()
    if not value:
        raise OutpatientScheduleSuggestionBridgeError(f"{key} is required")
    return value


def _parse_time(value: str):
    try:
        return datetime.strptime(value, "%H:%M").time()
    except ValueError as exc:
        raise OutpatientScheduleSuggestionBridgeError("time must use HH:MM") from exc


def _parse_mode(value: object) -> SuggestionMode:
    if isinstance(value, bool):
        raise OutpatientScheduleSuggestionBridgeError("mode is invalid")
    try:
        if isinstance(value, int):
            return SuggestionMode(value)
        text = str(value).strip()
        if text.isdigit():
            return SuggestionMode(int(text))
        return SuggestionMode(SUGGESTION_MODE_LABELS.index(text) + 1)
    except (ValueError, IndexError) as exc:
        raise OutpatientScheduleSuggestionBridgeError("mode is invalid") from exc


def run_outpatient_schedule_suggestion_bridge(
    payload: Mapping[str, object],
) -> dict[str, object]:
    source_path = str(payload.get("source_path", "")).strip()
    values = payload.get("values")
    if not source_path:
        raise OutpatientScheduleSuggestionBridgeError("source_path is required")
    if not isinstance(values, Mapping):
        raise OutpatientScheduleSuggestionBridgeError("values must be an object")

    source = Path(source_path)
    if not source.exists():
        raise OutpatientScheduleSuggestionBridgeError(f"source workbook not found: {source}")

    mode = _parse_mode(values.get("mode", ""))
    preferred_therapist = _required_text(values, "therapist")
    preferred_time = _parse_time(_required_text(values, "time"))
    preferred_days = _required_text(values, "days")
    limit_value = values.get("limit", 5)
    try:
        limit = int(limit_value)
    except (TypeError, ValueError) as exc:
        raise OutpatientScheduleSuggestionBridgeError("limit must be an integer") from exc
    limit = max(1, min(limit, 10))

    try:
        settings = read_settings(source)
        entries = read_unified_base_schedule(source)
        suggestions = suggest_outpatient_schedule_slots(
            mode=mode,
            preferred_therapist=preferred_therapist,
            preferred_time=preferred_time,
            preferred_day_pattern=preferred_days,
            therapist_names=settings.therapist_names,
            timeslots=settings.standard_timeslots,
            day_patterns=settings.day_patterns,
            existing_entries=entries,
            limit=limit,
        )
    except ValueError as exc:
        raise OutpatientScheduleSuggestionBridgeError(str(exc)) from exc

    return {
        "ok": True,
        "mode": int(mode),
        "mode_label": SUGGESTION_MODE_LABELS[int(mode) - 1],
        "suggestions": [
            {
                "therapist": item.therapist_id,
                "time": item.start_time.strftime("%H:%M"),
                "days": item.day_pattern,
            }
            for item in suggestions
        ],
    }
