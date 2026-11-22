"""Full private A:V vs admitted weekly rows reconciliation for the DEMO.

This is a read-only, fail-closed bridge. The review workbook is NOT accepted as
ground truth for treatment times or weekdays merely because its cell references
are correct. Every candidate is independently regenerated from the original
source and must match the admitted plan exactly, with full row multiplicity.

No patient-identifying data or original cell values are logged or printed.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
import re

from openpyxl import load_workbook

from rehab_excel.trial_admission_gate import (
    CONFIRMED_MARKERS, R6_DECISIONS_SHEET, SOURCE_WEEKLY_SHEET,
    _audit_r6_decision_readback, _data_rows, _r6_exact_refs,
)
from rehab_core.programme_mapping import (
    ProgrammeMappingError, SourceWeeklyEntry, WEEKDAYS, make_weekly_entries,
    parse_clock,
)
from rehab_excel.trial_import_plan import TrialImportPlan
from rehab_excel.trial_special_fields_audit import verify_legacy_afternoon_scope, verify_morning_special_scope, verify_untimed_legacy_scope
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



# Exact R6 operator-corrected *source field groups*, not a general override API.
# All other A:V groups continue through preview_source_weekdays unchanged.
_R6_CORRECTABLE = {
    "ΕΡΓΟΘΕΡΑΠΕΙΑ": ("Εργο", ("G", "H", "I")),
    "ΛΟΓΟΘΕΡΑΠΕΙΑ": ("Λογο", ("J", "K", "L")),
}
_R6_CORRECTED = "PROPOSED_OPERATOR_CORRECTION"
_R6_UNCHANGED = "PROPOSED_EXACT_SOURCE"


def _source_bound_r6_corrected_groups(review_book, src):
    """Read R6 operator changes as complete original-source treatment groups.

    No patient identity is inferred, no R6 status becomes clinical approval.
    The caller has already verified original/review hashes and source identity.
    Existing ledger after-values are checked by the existing admission audit.
    """
    if R6_DECISIONS_SHEET not in review_book.sheetnames:
        return {}
    if SOURCE_WEEKLY_SHEET not in review_book.sheetnames:
        raise TrialDemoSourceReconciliationError("R6 weekly sheet is missing")

    errors = []
    weekly = tuple(_data_rows(review_book[SOURCE_WEEKLY_SHEET], 11, blockers=errors))
    _audit_r6_decision_readback(review_book, weekly, errors)
    if errors:
        raise TrialDemoSourceReconciliationError(
            "R6 operator decision ledger does not match reviewed weekly rows"
        )
    ledger = review_book[R6_DECISIONS_SHEET]
    ledger_rows = tuple(_data_rows(ledger, 15, blockers=errors))
    if errors:
        raise TrialDemoSourceReconciliationError("R6 operator decision contains unsafe fields")

    corrected = defaultdict(list)
    for _, values in ledger_rows:
        if _text(values[14]) != _R6_CORRECTED:
            continue
        source_row, specialty = _source_row(values[2]), _text(values[4])
        if source_row is None or specialty not in _R6_CORRECTABLE:
            raise TrialDemoSourceReconciliationError("Unsupported R6 operator correction")
        corrected[(source_row, specialty)].append(values)
    if not corrected:
        return {}

    weekly_by_group = defaultdict(list)
    for _, values in weekly:
        number = _source_row(values[1])
        if (number, _text(values[3])) in corrected:
            weekly_by_group[(number, _text(values[3]))].append(values)

    result = {}
    for (source_row, specialty), decisions in corrected.items():
        treatment, columns = _R6_CORRECTABLE[specialty]
        if source_row > src.max_row:
            raise TrialDemoSourceReconciliationError("R6 corrected group is outside original A:V")
        expected_refs = frozenset((column, str(source_row)) for column in columns)
        original_cells = [
            src.cell(source_row, ord(column) - 64) for column in columns
        ]
        if any(cell.data_type in {"f", "e"} for cell in original_cells):
            raise TrialDemoSourceReconciliationError("R6 original treatment group is unsafe")
        original_clock, original_days, original_provider = [
            cell.value for cell in original_cells
        ]
        if not all(_text(v) for v in (original_clock, original_days, original_provider)):
            raise TrialDemoSourceReconciliationError("R6 corrected original field is incomplete")
        try:
            parsed_original_clock = parse_clock(original_clock)
        except (ProgrammeMappingError, TypeError, ValueError):
            raise TrialDemoSourceReconciliationError(
                "R6 corrected original clock is unreadable"
            ) from None
        before = []
        for decision in decisions:
            if frozenset(_r6_exact_refs(decision[9])) != expected_refs:
                raise TrialDemoSourceReconciliationError(
                    "R6 operator correction has wrong original cell references"
                )
            if _text(decision[10]):
                try:
                    old_clock = parse_clock(decision[6])
                except (ProgrammeMappingError, TypeError, ValueError):
                    old_clock = None
                if (
                    _text(decision[5]) not in WEEKDAYS
                    or old_clock != parsed_original_clock
                    or _text(decision[8]) != _text(original_provider)
                ):
                    raise TrialDemoSourceReconciliationError(
                        "R6 operator before-values differ from original source"
                    )
                before.append((
                    _text(decision[5]), old_clock, _text(decision[8]),
                ))
            elif "νέα ημέρα" not in _text(decision[0]).casefold():
                raise TrialDemoSourceReconciliationError(
                    "R6 new weekday lacks explicit operator decision"
                )
        if not before or len(before) != len(set(before)):
            raise TrialDemoSourceReconciliationError(
                "R6 original source baseline is missing or repeated"
            )

        # For parseable originals, all original weekday facts must appear in
        # the historical before-values. An extra weekday must be explicit.
        source_cell = f"{columns[0]}{source_row}"
        original_parsed = True
        try:
            baseline = make_weekly_entries(
                patient_id="FICTIONAL-SOURCE-COMPARISON", treatment=treatment,
                source_cell=source_cell, clock=original_clock,
                weekdays=original_days, therapist=original_provider,
            )
        except (ProgrammeMappingError, TypeError, ValueError):
            baseline = ()
            original_parsed = False
        if original_parsed and Counter(before) != Counter(
            (entry.weekday, entry.start_time, _text(entry.therapist_name))
            for entry in baseline
        ):
            raise TrialDemoSourceReconciliationError(
                "R6 operator baseline does not cover every original weekday"
            )

        final_group = weekly_by_group[(source_row, specialty)]
        if len(final_group) < len(decisions) or not final_group:
            raise TrialDemoSourceReconciliationError("R6 corrected group lost a final placement")
        final_rows = []
        seen_slots = set()
        for row in final_group:
            if (
                _text(row[10]) not in {_R6_CORRECTED, _R6_UNCHANGED}
                or frozenset(_r6_exact_refs(row[8])) != expected_refs
                or _text(row[4]) not in WEEKDAYS
                or _text(row[7]) == ""
            ):
                raise TrialDemoSourceReconciliationError(
                    "R6 corrected group has unapproved final source fields"
                )
            try:
                clock = parse_clock(row[5])
            except (ProgrammeMappingError, TypeError, ValueError):
                raise TrialDemoSourceReconciliationError(
                    "R6 corrected final weekday clock is unreadable"
                ) from None
            key = (_text(row[4]), clock)
            if key in seen_slots:
                raise TrialDemoSourceReconciliationError(
                    "R6 corrected group repeats the same weekday/clock"
                )
            seen_slots.add(key)
            final_rows.append((treatment, key[0], clock, _text(row[7]), source_cell))
            if _text(row[10]) == _R6_UNCHANGED:
                if original_parsed:
                    if (
                        key[0], clock, _text(row[7])
                    ) not in Counter(
                        (entry.weekday, entry.start_time, _text(entry.therapist_name))
                        for entry in baseline
                    ):
                        raise TrialDemoSourceReconciliationError(
                            "R6 unchanged group row differs from original source"
                        )
                elif (
                    clock != parsed_original_clock
                    or _text(row[7]) != _text(original_provider)
                ):
                    raise TrialDemoSourceReconciliationError(
                        "R6 unchanged ambiguous-source row differs from original clock/provider"
                    )

        if not original_parsed:
            # The legacy original has ambiguous joined day labels in this
            # case. Only the operator's explicitly written COMPLETE day list
            # may authorize preserving the source group; do not infer dates.
            explicit_day_sets = set()
            for decision in decisions:
                match = re.search(
                    r"(?:ΔΕ|ΤΡ|ΤΕ|ΠΕ|ΠΑ)(?:-(?:ΔΕ|ΤΡ|ΤΕ|ΠΕ|ΠΑ))+$",
                    _text(decision[0]),
                )
                if match:
                    explicit_day_sets.add(frozenset(match.group(0).split("-")))
            if (
                len(explicit_day_sets) != 1
                or next(iter(explicit_day_sets)) != frozenset(day for _, day, _, _, _ in final_rows)
                or len(final_rows) != len({day for _, day, _, _, _ in final_rows})
            ):
                raise TrialDemoSourceReconciliationError(
                    "R6 ambiguous original weekdays lack exact complete operator decision"
                )

        result[(source_row, columns[0])] = tuple(final_rows)
    return result


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

        src_book = load_workbook(programme, read_only=True, data_only=False)
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
                # Check coverage against the whole original A:V worksheet, not
                # merely the subset of patients appearing in the review copy.
                # A missing row with any C:V content is a hold, including
                # day/provider-only cells and orphan rows without a patient.
                # Formula/error cells must be resolved explicitly; a missing
                # cached formula value must never look like an empty treatment.
                active_source_rows: set[int] = set()
                for number, cells in enumerate(
                    src.iter_rows(min_row=2, max_col=SOURCE_WIDTH), start=2,
                ):
                    therapy_cells = cells[2:]
                    if any(cell.data_type in {"f", "e"} for cell in therapy_cells):
                        raise TrialDemoSourceReconciliationError(
                            "Original A:V formula or Excel error requires independent review"
                        )
                    if any(_text(cell.value) for cell in therapy_cells):
                        active_source_rows.add(number)

                try:
                    historical_rows = verify_legacy_afternoon_scope(programme=programme, review=review)
                except ValueError:
                    raise TrialDemoSourceReconciliationError(
                        "Historical afternoon scope requires independent review"
                    ) from None
                try:
                    special_days = verify_morning_special_scope(programme=programme, review=review)
                except ValueError:
                    raise TrialDemoSourceReconciliationError(
                        "Morning F/V scope requires independent review"
                    ) from None
                untimed_fields = verify_untimed_legacy_scope(programme=programme, review=review)
                corrected_groups = _source_bound_r6_corrected_groups(review_book, src)
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
                    # Only an exactly audited R6 source group may replace
                    # the original parser's ambiguous/moved group. The copy is
                    # in-memory; original A:V and review stay byte-identical.
                    preview_values = list(values)
                    selected_overrides = []
                    for (group_row, time_column), placements in corrected_groups.items():
                        if group_row != number:
                            continue
                        group_columns = next(
                            cols for _, cols in _R6_CORRECTABLE.values()
                            if cols[0] == time_column
                        )
                        for column in group_columns:
                            preview_values[ord(column) - ord("A")] = None
                        selected_overrides.extend(
                            SourceWeeklyEntry(pid, treatment, weekday, clock, provider, source_cell)
                            for treatment, weekday, clock, provider, source_cell in placements
                        )
                    proposed = preview_source_weekdays(
                        source_row=number, columns=tuple(preview_values),
                        confirmed_patient_id=pid,
                        historical_afternoon_acknowledged=number in historical_rows,
                        approved_morning_days=special_days,
                        acknowledged_untimed_fields=untimed_fields,
                    )
                    if proposed.blockers:
                        raise TrialDemoSourceReconciliationError(
                            "Original A:V treatment fields require independent review"
                        )
                    candidates.extend(proposed.entries)
                    candidates.extend(selected_overrides)

                if active_source_rows.difference(seen_source_rows):
                    raise TrialDemoSourceReconciliationError(
                        "Original A:V treatment rows are missing from approved review"
                    )
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
