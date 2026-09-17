from __future__ import annotations

from enum import IntEnum


class RehabWeekday(IntEnum):
    MONDAY = 0
    TUESDAY = 1
    WEDNESDAY = 2
    THURSDAY = 3
    FRIDAY = 4


_TOKEN_MAP = {
    # Monday
    "δ": RehabWeekday.MONDAY,
    "δε": RehabWeekday.MONDAY,
    "δευ": RehabWeekday.MONDAY,
    "δευτερα": RehabWeekday.MONDAY,
    # Tuesday
    "τρ": RehabWeekday.TUESDAY,
    "τρι": RehabWeekday.TUESDAY,
    "τριτη": RehabWeekday.TUESDAY,
    # Wednesday
    "τε": RehabWeekday.WEDNESDAY,
    "τετ": RehabWeekday.WEDNESDAY,
    "τεταρτη": RehabWeekday.WEDNESDAY,
    # Thursday
    "πε": RehabWeekday.THURSDAY,
    "πεμ": RehabWeekday.THURSDAY,
    "πεμπτη": RehabWeekday.THURSDAY,
    # Friday
    "π": RehabWeekday.FRIDAY,
    "πα": RehabWeekday.FRIDAY,
    "παρ": RehabWeekday.FRIDAY,
    "παρασκευη": RehabWeekday.FRIDAY,
}


def _normalize_token(token: str) -> str:
    return (
        token.strip()
        .casefold()
        .replace(".", "")
        .replace(" ", "")
        .replace("ά", "α")
        .replace("έ", "ε")
        .replace("ή", "η")
        .replace("ί", "ι")
        .replace("ό", "ο")
        .replace("ύ", "υ")
        .replace("ώ", "ω")
        .replace("ϊ", "ι")
        .replace("ΐ", "ι")
        .replace("ϋ", "υ")
        .replace("ΰ", "υ")
    )


def parse_day_pattern(pattern: str) -> frozenset[RehabWeekday]:
    """Parse the day abbreviations used by the rehab workbook.

    Examples include ``Δ-Τρ-Πε``, ``Δε-Τε-Πα`` and ``Δ-Τρ-Πε-Π``.
    Unknown tokens raise ValueError so invalid workbook data is not silently
    accepted.
    """

    if pattern is None or not pattern.strip():
        raise ValueError("Day pattern cannot be empty")

    normalized = pattern.replace("–", "-").replace("—", "-")
    tokens = [token for token in normalized.split("-") if token.strip()]
    if not tokens:
        raise ValueError("Day pattern cannot be empty")

    days: set[RehabWeekday] = set()
    for raw_token in tokens:
        token = _normalize_token(raw_token)
        try:
            days.add(_TOKEN_MAP[token])
        except KeyError as exc:
            raise ValueError(f"Unknown day token: {raw_token.strip()}") from exc

    return frozenset(days)


def patterns_overlap(first: str, second: str) -> bool:
    """Return True when two recurring patterns share at least one weekday."""

    return bool(parse_day_pattern(first) & parse_day_pattern(second))


def patterns_are_complementary(first: str, second: str) -> bool:
    """Return True when two valid recurring patterns do not share a weekday.

    This is the safe primitive needed later when deciding whether assignments
    can coexist because they apply on different days.
    """

    return not patterns_overlap(first, second)
