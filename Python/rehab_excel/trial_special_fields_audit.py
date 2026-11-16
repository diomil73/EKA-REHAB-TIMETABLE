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
from typing import Any

from openpyxl import load_workbook

from rehab_core.programme_mapping import ProgrammeMappingError, parse_clock
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
