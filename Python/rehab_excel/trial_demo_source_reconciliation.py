"""Full private A:V vs admitted weekly rows reconciliation for the DEMO.

This is a read-only, fail-closed bridge. The review workbook is NOT accepted as
ground truth for treatment times or weekdays merely because its cell references
are correct. Every candidate is independently regenerated from the original
source and must match the admitted plan exactly, with full row multiplicity.

No patient-identifying data or original cell values are logged or printed.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from rehab_excel.trial_admission_gate import CONFIRMED_MARKERS
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_source_alignment import (
    PATIENT_HEADERS, PATIENT_SHEET, PROGRAMME_HEADERS, PROGRAMME_SHEET,
    audit_trial_source_alignment,
)
from rehab_excel.trial_source_weekday_projection import (
    SOURCE_WIDTH, preview_source_weekdays,
)


class TrialDemoSourceReconciliationError(ValueError):
    """The original A:V source and approved weekly plan differ or are held."""


@dataclass(frozen=True)
class TrialDemoSourceReconciliation:
    patient_rows: int
    weekday_slots: int
    source_fields_held: int


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _source_row(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    try:
        n = int(value)
        return n if n > 1 and float(value) == n else None
    except (TypeError, ValueError, OverflowError):
        return None


def reconcile_demo_source_weekdays(
    plan: TrialImportPlan, *, programme: str | Path,
    review: str | Path, intake: str | Path,
) -> TrialDemoSourceReconciliation:
    """Require exact matching admitted entries for every confirmed patient.

    The source A:V data must be interpreted independently, preserving every
    source weekday, time, provider and source-cell origin. A value in any
    unsupported specialty (F/U/V) or any incomplete supported field causes
    a fail-closed stop, not a partial success.
    """
    if not isinstance(plan, TrialImportPlan) or plan.programme_sha256 is None:
        raise TrialDemoSourceReconciliationError(
            "Original A:V programme must be bound in the approved plan"
        )
    try:
        plan.verify_sources(review, intake, programme=programme)
        if not audit_trial_source_alignment(programme=programme, review=review).aligned:
            raise TrialDemoSourceReconciliationError(
                "Source row and cell reference alignment is not valid"
            )

        src_book = load_workbook(programme, read_only=True, data_only=True)
        try:
            review_book = load_workbook(review, read_only=True, data_only=True)
            try:
                src = src_book[PROGRAMME_SHEET]
                patients = review_book[PATIENT_SHEET]
                if (
                    tuple(src.cell(1, col).value for col in range(3, 23))
                    != PROGRAMME_HEADERS
                    or tuple(patients.cell(1, col).value for col in range(1, 7))
                    != PATIENT_HEADERS
                ):
                    raise TrialDemoSourceReconciliationError(
                        "Source or review header schema has changed"
                    )
                candidates = []
                seen_ids: set[str] = set()
                seen_source_rows: set[int] = set()
                patient_count = 0
                for row in patients.iter_rows(min_row=2, max_col=6, values_only=True):
                    if not any(_text(v) for v in row):
                        continue
                    original_row, clinic, name, context, pid, confirmation = row
                    number = _source_row(original_row)
                    pid = _text(pid)
                    if (
                        number is None or number > src.max_row
                        or not _text(clinic) or not _text(name) or not pid
                        or _text(confirmation).upper() not in CONFIRMED_MARKERS
                        or number in seen_source_rows or pid.casefold() in seen_ids
                    ):
                        raise TrialDemoSourceReconciliationError(
                            "Unconfirmed or repeated source patient identity"
                        )
                    seen_source_rows.add(number)
                    seen_ids.add(pid.casefold())
                    patient_count += 1
                    values = tuple(
                        src.cell(number, col).value
                        for col in range(1, SOURCE_WIDTH + 1)
                    )
                    if _text(values[0]) != _text(context) or _text(values[1]) != _text(name):
                        raise TrialDemoSourceReconciliationError(
                            "Source patient identity/context differs"
                        )
                    proposed = preview_source_weekdays(
                        source_row=number, columns=values,
                        confirmed_patient_id=pid,
                    )
                    if proposed.blockers:
                        raise TrialDemoSourceReconciliationError(
                            "Original A:V treatment fields require independent review"
                        )
                    candidates.extend(proposed.entries)

                if patient_count != plan.patient_count:
                    raise TrialDemoSourceReconciliationError(
                        "Admitted patient scope differs from original source"
                    )
                if (
                    len(candidates) != plan.weekday_slot_count
                    or Counter(candidates) != Counter(plan.entries)
                ):
                    raise TrialDemoSourceReconciliationError(
                        "Source weekday/time/provider/cell projection differs from approved plan"
                    )
            finally:
                review_book.close()
        finally:
            src_book.close()
        plan.verify_sources(review, intake, programme=programme)
    except TrialDemoSourceReconciliationError:
        raise
    except (OSError, ValueError, KeyError, TypeError, IndexError):
        raise TrialDemoSourceReconciliationError(
            "SAFETY STOP: missing, changed or unreadable original DEMO sources"
        ) from None

    return TrialDemoSourceReconciliation(
        patient_rows=patient_count,
        weekday_slots=len(candidates),
        source_fields_held=0,
    )
