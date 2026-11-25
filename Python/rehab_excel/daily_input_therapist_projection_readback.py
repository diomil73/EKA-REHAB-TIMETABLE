"""Read-only THERAPIST_DAILY occupied-clock projection diagnostic.

Reuses the two blocks, exact row/column positions and distinct-clock
interpretation from the current modTherapistWorkloadStrip VBA. This is NOT
a verifier of individual sessions, patient assignments, availability,
capacity, approved sharing, native recalculation or a journal transition.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time, timedelta
from hashlib import sha256
from pathlib import Path
import re
import unicodedata

from openpyxl import load_workbook

from rehab_core.recurring_conflicts import _provider_key

BLOCKS = ((1, 2, 8), (11, 12, 18))
PROVIDER_COLUMNS = tuple(range(2, 11))


class TherapistDailyProjectionReadbackError(RuntimeError):
    """Saved THERAPIST_DAILY occupied-clock projection is inconsistent."""


@dataclass(frozen=True)
class TherapistDailyOccupiedClockReadback:
    workbook_sha256: str
    checked_providers: int
    occupied_distinct_clocks: int


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as handle:
        for part in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(part)
    return h.hexdigest()


def _name(value) -> str:
    if type(value) is not str or not value.strip():
        raise TherapistDailyProjectionReadbackError("Provider label missing or invalid")
    key = _provider_key(value)
    if type(key) is not str or not key:
        raise TherapistDailyProjectionReadbackError("Provider identity invalid")
    return key


def _clock(value) -> time:
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    if isinstance(value, timedelta):
        minutes = round(value.total_seconds() / 60)
        if minutes < 0 or minutes >= 1440:
            raise TherapistDailyProjectionReadbackError("Time cell out of range")
        return time(minutes // 60, minutes % 60)
    if type(value) in (int, float):
        if value < 0 or value >= 1:
            raise TherapistDailyProjectionReadbackError("Time cell out of range")
        minutes = round(value * 1440) % 1440
        return time(minutes // 60, minutes % 60)
    if type(value) is str:
        raw = value.strip().replace(".", ":")
        if re.fullmatch(r"(?:[01]?\d|2[0-3]):[0-5]\d", raw):
            hour, minute = raw.split(":")
            return time(int(hour), int(minute))
    raise TherapistDailyProjectionReadbackError("Invalid therapist time slot")


def verify_therapist_daily_occupied_clocks(
    workbook_path: str | Path,
    *,
    expected_clocks: tuple[tuple[str, tuple[time, ...]], ...],
) -> TherapistDailyOccupiedClockReadback:
    """Compare independently expected provider occupied clocks to saved cells.

    Expectations must come from validated preflight facts, NEVER from this
    workbook's THERAPIST_DAILY sheet. A caller-selected scope could omit a
    provider: all non-empty provider headers are included and compared.
    """
    if type(expected_clocks) is not tuple or not expected_clocks:
        raise TherapistDailyProjectionReadbackError("Explicit occupied-clock expectations required")
    expected = {}
    for value in expected_clocks:
        if (type(value) is not tuple or len(value) != 2
                or type(value[1]) is not tuple
                or any(type(t) is not time for t in value[1])):
            raise TherapistDailyProjectionReadbackError("Malformed provider clock expectation")
        provider = _name(value[0])
        if provider in expected or len(set(value[1])) != len(value[1]):
            raise TherapistDailyProjectionReadbackError("Duplicate expected provider or clock")
        expected[provider] = frozenset(value[1])

    try:
        path = Path(workbook_path).resolve(strict=True)
        if not path.is_file() or path.suffix.casefold() != ".xlsm":
            raise TherapistDailyProjectionReadbackError("Saved XLSM required")
        before = _digest(path)
        wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    except TherapistDailyProjectionReadbackError:
        raise
    except Exception as exc:
        raise TherapistDailyProjectionReadbackError("Saved workbook unavailable") from exc

    try:
        if "THERAPIST_DAILY" not in wb.sheetnames:
            raise TherapistDailyProjectionReadbackError("THERAPIST_DAILY missing")
        ws = wb["THERAPIST_DAILY"]
        observed = {}
        for header, start, end in BLOCKS:
            for col in PROVIDER_COLUMNS:
                raw = ws.cell(header, col)
                if raw.data_type in {"f", "e"}:
                    raise TherapistDailyProjectionReadbackError("Provider header cannot be a formula/error")
                label = raw.value
                if label is None or (type(label) is str and not label.strip()):
                    # Data with no provider header must never be silently ignored.
                    if any(ws.cell(r, col).value not in (None, "") for r in range(start, end + 1)):
                        raise TherapistDailyProjectionReadbackError("Occupied rows under missing provider")
                    continue
                provider = _name(label)
                clocks = observed.setdefault(provider, set())
                for row in range(start, end + 1):
                    cell = ws.cell(row, col)
                    if cell.data_type == "e":
                        raise TherapistDailyProjectionReadbackError("Therapist view contains error")
                    # Formula results require real Excel recalculation; reject
                    # them rather than treating unevaluated expressions as occupancy.
                    if cell.data_type == "f":
                        raise TherapistDailyProjectionReadbackError("Formula occupancy requires native verification")
                    if cell.value is None or (type(cell.value) is str and not cell.value.strip()):
                        continue
                    clock_cell = ws.cell(row, 1)
                    if clock_cell.data_type in {"f", "e"}:
                        raise TherapistDailyProjectionReadbackError("Clock row contains formula/error")
                    clocks.add(_clock(clock_cell.value))
        if set(observed) != set(expected):
            raise TherapistDailyProjectionReadbackError("Provider coverage differs from preflight")
        if any(frozenset(observed[key]) != value for key, value in expected.items()):
            raise TherapistDailyProjectionReadbackError("Occupied clock membership differs from preflight")
        if _digest(path) != before:
            raise TherapistDailyProjectionReadbackError("Workbook changed during therapist view inspection")
        return TherapistDailyOccupiedClockReadback(
            workbook_sha256=before,
            checked_providers=len(observed),
            occupied_distinct_clocks=sum(len(x) for x in observed.values()),
        )
    except TherapistDailyProjectionReadbackError:
        raise
    except Exception as exc:
        raise TherapistDailyProjectionReadbackError("Therapist view readback invalid") from exc
    finally:
        wb.close()
