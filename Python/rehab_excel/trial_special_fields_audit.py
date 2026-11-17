"""Read-only audit of special A:V fields, with no patient values in results.

F: reclining/tilt-table, known 30-minute duration, but no approved weekdays.
U: legacy afternoon programme, explicitly NOT eligible for automatic seeding.
V: morning pool, exact weekday/provider/source semantics not yet approved.

This audit is a practical first gate on the actual manual source ALONE. It
never converts those fields into ordinary recurring slots. A field containing
an Excel formula is considered held even when Excel has no cached value.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook

from rehab_core.programme_mapping import ProgrammeMappingError, parse_clock, WEEKDAYS
from rehab_excel.trial_source_alignment import (
    PROGRAMME_HEADERS, PROGRAMME_SHEET, _headers_match,
)


class SpecialFieldsAuditError(ValueError):
    """Unsafe original programme or source changed during read-only inspection."""


@dataclass(frozen=True)
class SpecialFieldsAudit:
    patient_rows: int
    reclining_rows: int
    reclining_clock_parseable: int
    reclining_clock_review: int
    afternoon_rows: int
    pool_rows: int
    pool_clock_parseable: int
    pool_clock_review: int
    formula_fields: int
    orphan_rows: int
    source_unchanged: bool

    @property
    def unresolved_fields(self) -> int:
        """These rows remain HELD even if their clocks are machine-readable."""
        return self.reclining_rows + self.afternoon_rows + self.pool_rows

    @property
    def eligible_for_special_field_import(self) -> bool:
        """This module never authorizes any special-field import."""
        return False


def _digest(path: Path) -> str:
    value = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def _filled(value: Any) -> bool:
    return value is not None and bool(str(value).strip())


def _parseable_time(value: Any) -> bool:
    try:
        parse_clock(value)
        return True
    except (ProgrammeMappingError, ValueError, TypeError, OverflowError):
        return False


def inspect_real_manual_special_fields(path: str | Path) -> SpecialFieldsAudit:
    """Inspect current manual XLSX without patient-level data leaving memory."""
    source = Path(path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsx":
        raise SpecialFieldsAuditError("Original private .xlsx manual source required")

    before = _digest(source)
    patient_rows = 0
    f_rows = f_clocks = f_review = 0
    u_rows = 0
    v_rows = v_clocks = v_review = 0
    formulas = orphans = 0
    cached = None
    raw = None
    try:
        # A second raw-formula view is necessary because data_only=True can
        # suppress non-cached Excel formulas and silently lose source facts.
        cached = load_workbook(source, read_only=True, data_only=True)
        raw = load_workbook(source, read_only=True, data_only=False)
        if (
            PROGRAMME_SHEET not in cached.sheetnames
            or PROGRAMME_SHEET not in raw.sheetnames
        ):
            raise SpecialFieldsAuditError("Manual programme sheet missing")
        ws = cached[PROGRAMME_SHEET]
        ws_raw = raw[PROGRAMME_SHEET]
        if (
            not _headers_match(ws, PROGRAMME_HEADERS, start_col=3)
            or not _headers_match(ws_raw, PROGRAMME_HEADERS, start_col=3)
        ):
            raise SpecialFieldsAuditError("Manual A:V column contract differs")
        for row in range(2, max(ws.max_row, ws_raw.max_row) + 1):
            patient_present = (
                _filled(ws.cell(row, 2).value)
                or _filled(ws_raw.cell(row, 2).value)
            )
            if patient_present:
                patient_rows += 1

            present_columns = []
            for col in (6, 21, 22):
                observed = ws.cell(row, col).value
                original = ws_raw.cell(row, col)
                has_formula = original.data_type == "f"
                if not (_filled(observed) or _filled(original.value) or has_formula):
                    continue
                present_columns.append(col)
                if has_formula:
                    formulas += 1
                is_clock = not has_formula and _parseable_time(observed)
                if col == 6:
                    f_rows += 1
                    if is_clock:
                        f_clocks += 1
                    else:
                        f_review += 1
                elif col == 21:
                    # U is intentionally not treated as a scheduled session,
                    # regardless of whether it looks like a clock.
                    u_rows += 1
                else:
                    v_rows += 1
                    if is_clock:
                        v_clocks += 1
                    else:
                        v_review += 1
            if present_columns and not patient_present:
                orphans += 1
    except SpecialFieldsAuditError:
        raise
    except Exception:
        # No private raw source value, path or identifying exception text.
        raise SpecialFieldsAuditError("Cannot inspect private source safely") from None
    finally:
        if raw is not None:
            raw.close()
        if cached is not None:
            cached.close()
        if not source.is_file() or _digest(source) != before:
            raise SpecialFieldsAuditError(
                "SAFETY STOP: original manual source changed during inspection"
            )

    return SpecialFieldsAudit(
        patient_rows=patient_rows,
        reclining_rows=f_rows,
        reclining_clock_parseable=f_clocks,
        reclining_clock_review=f_review,
        afternoon_rows=u_rows,
        pool_rows=v_rows,
        pool_clock_parseable=v_clocks,
        pool_clock_review=v_review,
        formula_fields=formulas,
        orphan_rows=orphans,
        source_unchanged=True,
    )


AFTERNOON_SCOPE_SHEET = "ΕΓΚΡΙΣΕΙΣ_ΠΕΔΙΟΥ_U"
AFTERNOON_SCOPE_HEADERS = ("SourceCell", "ProgrammeSHA256", "Decision")
AFTERNOON_SCOPE_DECISION = "ACKNOWLEDGED_NOT_IMPORTED"


def verify_legacy_afternoon_scope(*, programme, review) -> frozenset[int]:
    """Verify an explicit, exact, source-bound non-import disposition of U.

    No afternoon appointments are created and no F/V holds are cleared.
    """
    p, r = Path(programme).resolve(), Path(review).resolve()
    if p == r or any(not x.is_file() or x.suffix.casefold() != ".xlsx" for x in (p, r)):
        raise SpecialFieldsAuditError("Separate original and review XLSX required")
    original = (_digest(p), _digest(r))
    source = review_book = None
    expected, approved = set(), set()
    try:
        source = load_workbook(p, read_only=True, data_only=False)
        review_book = load_workbook(r, read_only=True, data_only=False)
        ws = source[PROGRAMME_SHEET]
        if not _headers_match(ws, PROGRAMME_HEADERS, start_col=3):
            raise SpecialFieldsAuditError("Manual A:V column contract differs")
        for number, cells in enumerate(ws.iter_rows(min_row=2, max_col=21), start=2):
            patient, historical = cells[1], cells[20]
            if not _filled(historical.value):
                continue
            if (
                historical.data_type in {"f", "e"}
                or patient.data_type in {"f", "e"} or not _filled(patient.value)
            ):
                raise SpecialFieldsAuditError("Unsafe or orphaned historical afternoon cell")
            expected.add(number)
        if AFTERNOON_SCOPE_SHEET in review_book.sheetnames:
            scope = review_book[AFTERNOON_SCOPE_SHEET]
            if tuple(scope.cell(1, i).value for i in range(1, 4)) != AFTERNOON_SCOPE_HEADERS:
                raise SpecialFieldsAuditError("Historical afternoon scope headers differ")
            for cells in scope.iter_rows(min_row=2):
                if not any(_filled(cell.value) for cell in cells):
                    continue
                if (
                    len(cells) < 3 or any(cell.data_type in {"f", "e"} for cell in cells)
                    or any(_filled(cell.value) for cell in cells[3:])
                ):
                    raise SpecialFieldsAuditError("Historical afternoon scope must be literal")
                address, digest, decision = (cell.value for cell in cells[:3])
                if (
                    not isinstance(address, str)
                    or not re.fullmatch(r"U[1-9][0-9]*", address)
                    or digest != original[0] or decision != AFTERNOON_SCOPE_DECISION
                ):
                    raise SpecialFieldsAuditError("Invalid or stale historical afternoon acknowledgment")
                number = int(address[1:])
                if number in approved:
                    raise SpecialFieldsAuditError("Duplicate historical afternoon acknowledgment")
                approved.add(number)
        if expected != approved:
            raise SpecialFieldsAuditError("Historical afternoon cells require exact scope acknowledgment")
    except SpecialFieldsAuditError:
        raise
    except Exception:
        raise SpecialFieldsAuditError("Cannot verify historical afternoon scope safely") from None
    finally:
        if source is not None:
            source.close()
        if review_book is not None:
            review_book.close()
        if (_digest(p), _digest(r)) != original:
            raise SpecialFieldsAuditError("Source changed during historical afternoon scope verification")
    return frozenset(approved)


MORNING_SPECIAL_SHEET = "ΕΓΚΡΙΣΕΙΣ_ΠΕΔΙΩΝ_FV"
MORNING_SPECIAL_HEADERS = ("SourceCell", "ProgrammeSHA256", "Weekdays", "Decision")
MORNING_SPECIAL_DECISION = "APPROVED_MORNING_SERVICE_SLOT"


def verify_morning_special_scope(*, programme, review):
    """Verify exact F/V source-cell weekday and service approval.

    The authorized reviewer confirms actual service/resource availability and
    full interval compatibility. Preserve the original clock, never infer a
    named provider for the existing providerless F/V planner fields.
    """
    p, r = Path(programme).resolve(), Path(review).resolve()
    if p == r or any(not x.is_file() or x.suffix.casefold() != ".xlsx" for x in (p, r)):
        raise SpecialFieldsAuditError("Separate original and review XLSX required")
    original = (_digest(p), _digest(r))
    source = staged = None
    expected, approved = set(), {}
    try:
        source = load_workbook(p, read_only=True, data_only=False)
        staged = load_workbook(r, read_only=True, data_only=False)
        ws = source[PROGRAMME_SHEET]
        if not _headers_match(ws, PROGRAMME_HEADERS, start_col=3):
            raise SpecialFieldsAuditError("Manual A:V column contract differs")
        for number, cells in enumerate(ws.iter_rows(min_row=2, max_col=22), start=2):
            for letter, col in (("F", 6), ("V", 22)):
                cell = cells[col - 1]
                if not _filled(cell.value):
                    continue
                if (cell.data_type in {"f", "e"} or cells[1].data_type in {"f", "e"}
                        or not _filled(cells[1].value)):
                    raise SpecialFieldsAuditError("Unsafe or orphaned morning special cell")
                clock = parse_clock(cell.value)
                if letter == "F" and clock.hour * 60 + clock.minute + 30 > 1440:
                    raise SpecialFieldsAuditError("Reclining interval exceeds day")
                expected.add(f"{letter}{number}")
        if MORNING_SPECIAL_SHEET in staged.sheetnames:
            ws = staged[MORNING_SPECIAL_SHEET]
            if tuple(ws.cell(1, col).value for col in range(1, 5)) != MORNING_SPECIAL_HEADERS:
                raise SpecialFieldsAuditError("Morning special approval header mismatch")
            for cells in ws.iter_rows(min_row=2):
                if not any(_filled(cell.value) for cell in cells):
                    continue
                if (len(cells) < 4 or any(cell.data_type in {"f", "e"} for cell in cells)
                        or any(_filled(cell.value) for cell in cells[4:])):
                    raise SpecialFieldsAuditError("Morning special approval must be literal")
                address, digest, day_text, decision = (cell.value for cell in cells[:4])
                days = tuple(day_text.split("-")) if isinstance(day_text, str) else ()
                if (not isinstance(address, str) or address not in expected
                        or address in approved or digest != original[0]
                        or decision != MORNING_SPECIAL_DECISION
                        or not days or any(day not in WEEKDAYS for day in days)
                        or len(set(days)) != len(days)):
                    raise SpecialFieldsAuditError("Invalid or stale morning special approval")
                approved[address] = tuple(day for day in WEEKDAYS if day in days)
        if approved.keys() != expected:
            raise SpecialFieldsAuditError("Morning F/V cells require exact independent review")
    except SpecialFieldsAuditError:
        raise
    except Exception:
        raise SpecialFieldsAuditError("Cannot verify private morning special scope") from None
    finally:
        if source is not None:
            source.close()
        if staged is not None:
            staged.close()
        if (_digest(p), _digest(r)) != original:
            raise SpecialFieldsAuditError("Source changed during morning special scope verification")
    return approved

UNTIMED_SCOPE_SHEET = "ΕΓΚΡΙΣΕΙΣ_ΧΩΡΙΣ_ΩΡΑ"
UNTIMED_SCOPE_HEADERS = ("SourceCell", "ProgrammeSHA256", "Decision")


def verify_untimed_legacy_scope(*, programme, review) -> frozenset[str]:
    """Recognize only explicitly acknowledged occupational marks and untimed psychology.

    Acknowledgment is source-bound, never a clinical appointment or an
    authorization to change any original field. Missing decisions remain
    blockers in the weekday projection.
    """
    p, r = Path(programme).resolve(), Path(review).resolve()
    if p == r or any(not x.is_file() or x.suffix.casefold() != ".xlsx" for x in (p, r)):
        raise SpecialFieldsAuditError("Separate original and review XLSX required")
    original = (_digest(p), _digest(r))
    source = staged = None
    allowed, approved = {}, set()
    try:
        source = load_workbook(p, read_only=True, data_only=False)
        staged = load_workbook(r, read_only=True, data_only=False)
        ws = source[PROGRAMME_SHEET]
        if not _headers_match(ws, PROGRAMME_HEADERS, start_col=3):
            raise SpecialFieldsAuditError("Manual A:V column contract differs")
        for number, cells in enumerate(ws.iter_rows(min_row=2, max_col=15), start=2):
            if not _filled(cells[1].value) or cells[1].data_type in {"f", "e"}:
                continue
            if (cells[6].value == "*" and
                    all(c.data_type not in {"f", "e"} for c in cells[6:9])):
                allowed[f"G{number}"] = "IGNORE_OCCUPATIONAL_MARK"
            if (not _filled(cells[12].value)
                    and any(_filled(c.value) for c in cells[13:15])
                    and all(c.data_type not in {"f", "e"} for c in cells[12:15])):
                allowed[f"M{number}"] = "PSYCHOLOGY_TIME_PENDING"
        if UNTIMED_SCOPE_SHEET in staged.sheetnames:
            scope = staged[UNTIMED_SCOPE_SHEET]
            if tuple(scope.cell(1, i).value for i in range(1, 4)) != UNTIMED_SCOPE_HEADERS:
                raise SpecialFieldsAuditError("Untimed scope header mismatch")
            for cells in scope.iter_rows(min_row=2, min_col=1, max_col=max(3, scope.max_column or 3)):
                if not any(_filled(c.value) or c.data_type in {"f", "e"} for c in cells):
                    continue
                if (any(c.data_type in {"f", "e"} for c in cells)
                        or any(_filled(c.value) for c in cells[3:])):
                    raise SpecialFieldsAuditError("Untimed scope must be literal")
                address, digest, decision = (c.value for c in cells[:3])
                if (not isinstance(address, str) or address in approved
                        or digest != original[0] or address not in allowed
                        or decision != allowed[address]):
                    raise SpecialFieldsAuditError("Invalid or stale untimed disposition")
                approved.add(address)
    except SpecialFieldsAuditError:
        raise
    except Exception:
        raise SpecialFieldsAuditError("Cannot verify private untimed scope") from None
    finally:
        if source is not None:
            source.close()
        if staged is not None:
            staged.close()
        if (_digest(p), _digest(r)) != original:
            raise SpecialFieldsAuditError("Source changed during untimed scope verification")
    return frozenset(approved)
