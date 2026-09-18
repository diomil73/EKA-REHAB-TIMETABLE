from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from pathlib import Path
import re
import unicodedata
from typing import Iterable, Mapping

from .package_probe import WorkbookPackageProbe


class LayoutSafetyError(RuntimeError):
    """Raised when a workbook cell is outside the confirmed layout contract."""


@dataclass(frozen=True)
class CellRange:
    sheet: str
    start: str
    end: str
    purpose: str


@dataclass(frozen=True)
class LayoutIssue:
    code: str
    message: str
    severity: str = "error"


@dataclass(frozen=True)
class LayoutAudit:
    workbook_path: str
    issues: tuple[LayoutIssue, ...]

    @property
    def ok(self) -> bool:
        return not any(issue.severity == "error" for issue in self.issues)


MASTER_SCHEDULE_HEADERS: Mapping[str, str] = {
    "A1": "Μολυσματικός",
    "B1": "Θάλαμος",
    "C1": "Ασθενής",
    "D1": "ΦΘ",
    "E1": "Ρομποτικό",
    "F1": "Πισίνα",
    "G1": "Ανακλινόμενο",
    "H1": "Εργο",
    "I1": "Λογο",
    "J1": "ΕΦΑ",
    "K1": "Κατάσταση",
}

THERAPIST_DAILY_HEADER_RANGES = ("B1:J1", "B11:J11")
THERAPIST_DAILY_GRID_RANGES = ("B2:J8", "B12:J18")
THERAPIST_DAILY_TIMES = (
    time(8, 30),
    time(9, 15),
    time(10, 0),
    time(10, 45),
    time(11, 30),
    time(12, 15),
    time(13, 0),
)

THERAPIST_DAILY_BLOCKS = (
    # header_row, first_time_row, last_time_row
    (1, 2, 8),
    (11, 12, 18),
)

# Confirmed current cell-level write zones. A sheet being a future output sheet
# is not sufficient: a cell must also fall inside one of these mapped zones.
WRITE_ZONES: tuple[CellRange, ...] = (
    CellRange(
        "THERAPIST_DAILY",
        "B2",
        "J8",
        "Top therapist daily patient grid; headers/times/button excluded",
    ),
    CellRange(
        "THERAPIST_DAILY",
        "B12",
        "J18",
        "Bottom therapist daily patient grid; headers/times excluded",
    ),
    CellRange(
        "MASTER_SCHEDULE",
        "D2",
        "J5000",
        "Derived therapy display cells only; identity/status columns protected",
    ),
    CellRange(
        "CONFLICT_LOG",
        "A2",
        "A5000",
        "Conflict output rows; title A1 protected",
    ),
    CellRange(
        "REPLACEMENTS",
        "A1",
        "V5000",
        "Legacy dynamic replacement canvas; drawing buttons handled separately",
    ),
)

REPLACEMENT_LOG_HEADERS: Mapping[str, str] = {
    "A1": "Timestamp",
    "B1": "Ασθενής",
    "C1": "Αρχικός Θεραπευτής",
    "D1": "Αρχική Ώρα",
    "E1": "Νεος Θεραπευτής",
    "F1": "Νέα Ώρα",
    "G1": "Status",
    "H1": "Note",
}

THERAPIST_ATTENDANCE_HEADERS: Mapping[str, str] = {
    "A1": "Θεραπευτής",
    "B1": "Παρουσία (Ν/Ο)",
    "C1": "Slots φορτίου",
    "D1": "Καταχωρήσεις",
    "E1": "Μολυσματικοί",
    "F1": "Σύγκρουση",
}


_CELL_RE = re.compile(r"^([A-Z]+)([1-9]\d*)$")


def _column_number(column: str) -> int:
    number = 0
    for char in column.upper():
        if not ("A" <= char <= "Z"):
            raise LayoutSafetyError(f"Invalid Excel column: {column!r}")
        number = number * 26 + (ord(char) - ord("A") + 1)
    return number


def _column_letters(number: int) -> str:
    if number < 1:
        raise LayoutSafetyError(f"Invalid Excel column number: {number}")
    chars: list[str] = []
    while number:
        number, rem = divmod(number - 1, 26)
        chars.append(chr(ord("A") + rem))
    return "".join(reversed(chars))


def parse_cell(cell: str) -> tuple[int, int]:
    match = _CELL_RE.fullmatch(cell.strip().upper())
    if not match:
        raise LayoutSafetyError(f"Invalid Excel cell reference: {cell!r}")
    return int(match.group(2)), _column_number(match.group(1))


def cell_address(row: int, column: int) -> str:
    if row < 1:
        raise LayoutSafetyError(f"Invalid Excel row: {row}")
    return f"{_column_letters(column)}{row}"


def cell_in_range(cell: str, start: str, end: str) -> bool:
    row, column = parse_cell(cell)
    start_row, start_col = parse_cell(start)
    end_row, end_col = parse_cell(end)
    return start_row <= row <= end_row and start_col <= column <= end_col


def write_zones_for_sheet(sheet: str) -> tuple[CellRange, ...]:
    return tuple(zone for zone in WRITE_ZONES if zone.sheet == sheet)


def assert_cell_write_allowed(sheet: str, cell: str) -> None:
    zones = write_zones_for_sheet(sheet)
    if not zones:
        raise LayoutSafetyError(
            f"Sheet {sheet} has no confirmed cell-level write zone yet"
        )
    if not any(cell_in_range(cell, zone.start, zone.end) for zone in zones):
        ranges = ", ".join(f"{zone.start}:{zone.end}" for zone in zones)
        raise LayoutSafetyError(
            f"Cell {sheet}!{cell.upper()} is outside confirmed write zones: {ranges}"
        )


def _clean_text(value: object) -> str:
    return str(value).strip() if value is not None else ""


def _name_key(value: object) -> str:
    text = " ".join(_clean_text(value).casefold().split())
    # Excel names are not consistently accented/cased across legacy sheets.
    # Remove combining marks so ΜΗΛΙΔΑΚΗΣ and Μηλιδάκης resolve identically.
    decomposed = unicodedata.normalize("NFD", text)
    return "".join(char for char in decomposed if not unicodedata.combining(char))


def _excel_time(value: object) -> time | None:
    if value is None:
        return None
    if isinstance(value, time):
        return value.replace(second=0, microsecond=0)
    text = _clean_text(value)
    if not text:
        return None
    try:
        fraction = float(text) % 1
    except ValueError:
        normalized = text.replace(".", ":")
        match = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", normalized)
        if not match:
            return None
        return time(int(match.group(1)), int(match.group(2)))
    minutes = round(fraction * 24 * 60)
    return time((minutes // 60) % 24, minutes % 60)


def resolve_therapist_daily_cell_from_cells(
    cells: Mapping[str, object],
    therapist_name: str,
    slot_time: time,
) -> str:
    target_name = _name_key(therapist_name)
    target_time = slot_time.replace(second=0, microsecond=0)

    matches: list[str] = []
    for header_row, first_time_row, last_time_row in THERAPIST_DAILY_BLOCKS:
        provider_col: int | None = None
        for col in range(2, 11):  # B:J
            if _name_key(cells.get(cell_address(header_row, col))) == target_name:
                provider_col = col
                break
        if provider_col is None:
            continue

        for row in range(first_time_row, last_time_row + 1):
            if _excel_time(cells.get(cell_address(row, 1))) == target_time:
                matches.append(cell_address(row, provider_col))

    if not matches:
        raise LayoutSafetyError(
            f"No THERAPIST_DAILY cell found for {therapist_name!r} at {target_time:%H:%M}"
        )
    if len(matches) > 1:
        raise LayoutSafetyError(
            f"Ambiguous THERAPIST_DAILY mapping for {therapist_name!r} at "
            f"{target_time:%H:%M}: {matches}"
        )
    return matches[0]


def resolve_therapist_daily_cell(
    workbook_path: str | Path,
    therapist_name: str,
    slot_time: time,
) -> str:
    probe = WorkbookPackageProbe(workbook_path)
    return resolve_therapist_daily_cell_from_cells(
        probe.cell_values("THERAPIST_DAILY"),
        therapist_name,
        slot_time,
    )


def resolve_master_schedule_rows_from_cells(
    cells: Mapping[str, object],
    patient_name: str,
    *,
    first_row: int = 2,
    last_row: int = 5000,
) -> tuple[int, ...]:
    target = _name_key(patient_name)
    return tuple(
        row
        for row in range(first_row, last_row + 1)
        if _name_key(cells.get(f"C{row}")) == target
    )


def resolve_master_schedule_rows(
    workbook_path: str | Path,
    patient_name: str,
) -> tuple[int, ...]:
    probe = WorkbookPackageProbe(workbook_path)
    return resolve_master_schedule_rows_from_cells(
        probe.cell_values("MASTER_SCHEDULE"), patient_name
    )


def _expect_cells(
    cells: Mapping[str, object],
    expected: Mapping[str, str],
    *,
    code_prefix: str,
) -> list[LayoutIssue]:
    issues: list[LayoutIssue] = []
    for address, expected_value in expected.items():
        actual = _clean_text(cells.get(address))
        if actual != expected_value:
            issues.append(
                LayoutIssue(
                    f"{code_prefix}_CELL_MISMATCH",
                    f"{address}: expected {expected_value!r}, found {actual!r}",
                )
            )
    return issues


def audit_layout(workbook_path: str | Path) -> LayoutAudit:
    probe = WorkbookPackageProbe(workbook_path)
    issues: list[LayoutIssue] = []

    required = {
        "MASTER_SCHEDULE",
        "THERAPIST_DAILY",
        "REPLACEMENTS",
        "REPLACEMENT_LOG",
        "CONFLICT_LOG",
        "THERAPIST_ATTENDANCE",
    }
    missing = sorted(required.difference(probe.sheet_names))
    for sheet in missing:
        issues.append(LayoutIssue("MISSING_LAYOUT_SHEET", f"Missing sheet: {sheet}"))
    if missing:
        return LayoutAudit(str(workbook_path), tuple(issues))

    master = probe.cell_values("MASTER_SCHEDULE")
    issues.extend(
        _expect_cells(master, MASTER_SCHEDULE_HEADERS, code_prefix="MASTER_SCHEDULE")
    )

    daily = probe.cell_values("THERAPIST_DAILY")
    for header_row, first_time_row, last_time_row in THERAPIST_DAILY_BLOCKS:
        nonempty_headers = [
            daily.get(cell_address(header_row, col)) for col in range(2, 11)
        ]
        if any(not _clean_text(value) for value in nonempty_headers):
            issues.append(
                LayoutIssue(
                    "THERAPIST_DAILY_HEADER_GAP",
                    f"Expected populated therapist headers in B{header_row}:J{header_row}",
                )
            )
        actual_times = tuple(
            _excel_time(daily.get(cell_address(row, 1)))
            for row in range(first_time_row, last_time_row + 1)
        )
        if actual_times != THERAPIST_DAILY_TIMES:
            issues.append(
                LayoutIssue(
                    "THERAPIST_DAILY_TIME_GRID_MISMATCH",
                    f"Unexpected time grid in A{first_time_row}:A{last_time_row}: "
                    f"{actual_times}",
                )
            )

    replacement_log = probe.cell_values("REPLACEMENT_LOG")
    issues.extend(
        _expect_cells(
            replacement_log,
            REPLACEMENT_LOG_HEADERS,
            code_prefix="REPLACEMENT_LOG",
        )
    )

    attendance = probe.cell_values("THERAPIST_ATTENDANCE")
    issues.extend(
        _expect_cells(
            attendance,
            THERAPIST_ATTENDANCE_HEADERS,
            code_prefix="THERAPIST_ATTENDANCE",
        )
    )

    conflicts = probe.cell_values("CONFLICT_LOG")
    if _clean_text(conflicts.get("A1")) != "Συγκρούσεις / Υπερβάσεις":
        issues.append(
            LayoutIssue(
                "CONFLICT_LOG_TITLE_MISMATCH",
                "CONFLICT_LOG!A1 no longer contains the confirmed title",
            )
        )

    refresh_macros = probe.legacy_button_macros("THERAPIST_DAILY")
    if "RefreshAll" not in refresh_macros:
        issues.append(
            LayoutIssue(
                "THERAPIST_DAILY_REFRESH_BUTTON_CHANGED",
                "Existing RefreshAll form-control macro was not found",
                severity="warning",
            )
        )

    replacement_macros = probe.drawing_macros("REPLACEMENTS")
    if "ApplyReplacementChoiceV4" not in replacement_macros:
        issues.append(
            LayoutIssue(
                "REPLACEMENTS_BUTTON_MACRO_CHANGED",
                "Existing ApplyReplacementChoiceV4 drawing buttons were not found",
                severity="warning",
            )
        )

    return LayoutAudit(str(workbook_path), tuple(issues))
