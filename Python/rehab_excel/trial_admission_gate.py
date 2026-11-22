"""Fail-closed admission gate for the isolated real-data Diomil trial.

Consumes the local-only reviewed staging workbook, NEVER uploads identities,
writes an .xlsm, assigns IDs by guessing, or silently drops unresolved cells.
A positive gate means the workbook is eligible for a separately audited
native-Windows Excel trial import; it is NOT evidence of an import.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import Any

from openpyxl import load_workbook
from rehab_core.recurring_conflicts import _provider_key

from rehab_excel.trial_source_alignment import (
    TREATMENT_SOURCE_COLUMNS, PROGRAMME_SHEET, audit_trial_source_alignment,
)
from rehab_core.programme_mapping import (
    ProgrammeMappingError,
    SourceWeeklyEntry,
    WEEKDAYS,
    parse_clock, make_weekly_entries,
    validate_weekly_entry_consistency,
)

SOURCE_PATIENT_SHEET = "ΑΣΘΕΝΕΙΣ"
SOURCE_WEEKLY_SHEET = "ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"
SOURCE_HOLDS_SHEET = "ΕΓΓΡΑΦΕΣ_ΠΡΟΣ_ΕΛΕΓΧΟ"
SOURCE_OVERLAP_SHEET = "ΕΠΙΚΑΛΥΨΕΙΣ"
SOURCE_SHARED_SHEET = "ΕΓΚΡΙΣΕΙΣ_ΚΟΙΝΩΝ_SLOTS"
SHARED_HEADERS = ("Therapist", "Weekday", "StartTime", "PatientIDs", "Decision")
SHARED_ACCEPTED = "APPROVED_SHARED"
REQUIRED_SHEETS = frozenset({
    SOURCE_PATIENT_SHEET, SOURCE_WEEKLY_SHEET,
    SOURCE_HOLDS_SHEET, SOURCE_OVERLAP_SHEET,
})
CONFIRMED_MARKERS = frozenset({"CONFIRMED", "ΕΠΙΒΕΒΑΙΩΜΕΝΟ", "ΕΠΙΒΕΒΑΙΩΘΗΚΕ"})
SOURCE_TREATMENTS = {
    "ΦΘ": "ΦΘ",
    "ΡΟΜΠΟΤΙΚΗ": "Ρομποτικό",
    "ΕΡΓΟΘΕΡΑΠΕΙΑ": "Εργο",
    "ΛΟΓΟΘΕΡΑΠΕΙΑ": "Λογο",
    "ΨΥΧΟΛΟΓΙΑ": "Ψυχ",
    "ΕΦΑ": "ΕΦΑ",
    "ΑΝΑΚΛΙΝΟΜΕΝΟ": "Ανακλινόμενο",
    "ΠΙΣΙΝΑ": "Πισίνα",
}
# The original programme has one global row index, irrespective of clinic.
# Refuse non-time cells, references to another patient row and duplicate refs.
SOURCE_TIME_COLUMNS = {
    "ΦΘ": "C",
    "ΑΝΑΚΛΙΝΟΜΕΝΟ": "F",
    "ΕΡΓΟΘΕΡΑΠΕΙΑ": "G",
    "ΛΟΓΟΘΕΡΑΠΕΙΑ": "J",
    "ΨΥΧΟΛΟΓΙΑ": "M",
    "ΕΦΑ": "P",
    "ΡΟΜΠΟΤΙΚΗ": "R",
    "ΠΙΣΙΝΑ": "V",
}
SOURCE_REF_COLUMNS = {
    **TREATMENT_SOURCE_COLUMNS,
    "ΠΙΣΙΝΑ": frozenset({"V"}),
}
SRC_CELL_RE = re.compile(r"\b([A-Z]{1,3})([1-9][0-9]*)\b")


class TrialAdmissionError(ValueError):
    """Source workbook incomplete or not independently reconciled."""


@dataclass(frozen=True)
class TrialAdmissionResult:
    candidate_patients: int
    verified_patient_ids: int
    weekly_proposals: int
    unadmitted_weekly_rows: int
    unresolved_source_rows: int
    overlap_rows: int
    duplicate_ids: int
    admitted_entries: tuple[SourceWeeklyEntry, ...] = field(repr=False)
    blockers: tuple[str, ...]
    confirmed_combined_slots: frozenset[frozenset[SourceWeeklyEntry]] = field(
        default_factory=frozenset, repr=False,
    )

    @property
    def ready(self) -> bool:
        return not self.blockers


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _source_row_number(value: Any) -> int | None:
    """Accept only a positive integral original Excel row, not guessed IDs."""
    if isinstance(value, bool):
        return None
    try:
        number = int(value)
        return number if number >= 2 and float(value) == number else None
    except (TypeError, ValueError, OverflowError):
        return None


def _data_rows(ws, width: int, *, blockers: list[str]):
    """Count raw rows, even if a formula has no cached/displayed value.

    Review approvals are explicit literal values. Neither formulas nor Excel
    error cells may serve as clinical decisions or make held rows disappear.
    Mask unsafe cell content before parsing, keeping the row for diagnostics.
    """
    for number, cells in enumerate(
        ws.iter_rows(min_row=2, max_col=width), start=2
    ):
        unsafe = tuple(cell.data_type in {"f", "e"} for cell in cells)
        if any(unsafe):
            blockers.append(f"Review row {number}: formula or Excel error in admission fields")
        if any(unsafe) or any(_text(cell.value) for cell in cells):
            yield number, tuple(
                None if invalid else cell.value
                for cell, invalid in zip(cells, unsafe)
            )



R6_DECISIONS_SHEET = "ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ"
R6_DECISION_HEADERS = (
    "Απόφαση", "Γραμμή R5", "Γραμμή πηγής", "Ασθενής",
    "Ειδικότητα", "Πριν ημέρα", "Πριν ώρα", "Πριν ΦΘ/Ρομπ.",
    "Πριν όνομα", "Κελιά", "Πριν κατάσταση", "Μετά ημέρα",
    "Μετά ώρα", "Μετά όνομα", "Μετά κατάσταση",
)
R6_LEDGER_RETAINED = frozenset({
    "PROPOSED_OPERATOR_CORRECTION", "HOLD_APPROVED_COMBINED",
    "HOLD_RESCHEDULE",
})
R6_LEDGER_EXCLUDED = "EXCLUDED_FROM_DRAFT"


def _r6_exact_refs(raw: object) -> tuple[tuple[str, str], ...]:
    """Accept complete, explicit source cell addresses, not a fuzzy regex match."""
    tokens = [part.strip() for part in _text(raw).upper().split("/")]
    if not tokens or any(
        re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", token) is None
        for token in tokens
    ):
        return ()
    return tuple(SRC_CELL_RE.findall(token)[0] for token in tokens)


def _audit_r6_decision_readback(wb, proposals, blockers: list[str]) -> None:
    """Cross-check recorded R6 after-values against the final weekly view.

    This checks ledger completeness and multiplicity ONLY. It neither approves
    corrected clinical slots nor permits held/combined entries to be imported.
    Original A:V source interpretation remains an independent mandatory gate.
    """
    if R6_DECISIONS_SHEET not in wb.sheetnames:
        return  # Backwards compatibility: non-R6 reviews retain strict admission.

    ledger = wb[R6_DECISIONS_SHEET]
    if tuple(ledger.cell(1, i).value for i in range(1, 16)) != R6_DECISION_HEADERS:
        blockers.append("R6 decision ledger headers differ from the approved schema")
        return

    def key(source, name, specialty, weekday, clock, refs, state):
        source_num = _source_row_number(source)
        refs_exact = _r6_exact_refs(refs)
        if source_num is None or not refs_exact:
            return None
        try:
            normalized_clock = parse_clock(clock)
        except (ProgrammeMappingError, TypeError, ValueError):
            return None
        return (
            source_num, _text(name), _text(specialty), _text(weekday),
            normalized_clock, refs_exact, _text(state),
        )

    weekly: dict[tuple, list[tuple[int, tuple]]] = {}
    controlled_weekly_rows: set[int] = set()
    for number, values in proposals:
        state = _text(values[10])
        if state not in R6_LEDGER_RETAINED:
            continue
        controlled_weekly_rows.add(number)
        item_key = key(
            values[1], values[2], values[3], values[4],
            values[5], values[8], state,
        )
        if item_key is None:
            blockers.append(f"R6 weekly decision row {number}: invalid source provenance")
            continue
        weekly.setdefault(item_key, []).append((number, values))

    seen_r5: set[int] = set()
    matched: set[int] = set()
    for number, values in _data_rows(ledger, 15, blockers=blockers):
        (
            decision, r5_row, source, name, specialty,
            before_day, before_clock, before_physio, before_optional,
            refs, before_state, after_day, after_clock, after_name, state,
        ) = values
        state = _text(state)
        r5 = _source_row_number(r5_row)
        if (
            not _text(decision) or r5 is None or r5 in seen_r5
            or _source_row_number(source) is None or not _text(name)
            or not _text(specialty) or not _r6_exact_refs(refs)
            or not _text(before_day)
            or not (
                _text(before_state)
                or (
                    state == "PROPOSED_OPERATOR_CORRECTION"
                    and "νέα ημέρα" in _text(decision).casefold()
                )
            )
        ):
            blockers.append(f"R6 decision ledger row {number}: incomplete or repeated provenance")
            continue
        seen_r5.add(r5)
        if state == R6_LEDGER_EXCLUDED:
            if any(_text(v) for v in (after_day, after_clock, after_name)):
                blockers.append(f"R6 decision ledger row {number}: excluded entry has after-values")
            continue
        if state not in R6_LEDGER_RETAINED or not _text(after_day):
            blockers.append(f"R6 decision ledger row {number}: unsupported disposition")
            continue
        after_key = key(
            source, name, specialty, after_day, after_clock, refs, state,
        )
        matching = weekly.get(after_key, []) if after_key is not None else []
        if len(matching) != 1 or matching[0][0] in matched:
            blockers.append(f"R6 decision ledger row {number}: after-values not unique in weekly review")
            continue
        weekly_row, final_values = matching[0]
        actual_provider = (
            final_values[6] if _text(specialty) in {"ΦΘ", "ΡΟΜΠΟΤΙΚΗ"}
            else final_values[7]
        )
        if _text(after_name) != _text(actual_provider):
            blockers.append(f"R6 decision ledger row {number}: after-provider differs")
            continue
        if state == "HOLD_APPROVED_COMBINED":
            # An approved combined slot is a preserved original placement,
            # not permission to shift its day, time or therapist.
            try:
                unchanged_time = parse_clock(before_clock) == parse_clock(after_clock)
            except (ProgrammeMappingError, TypeError, ValueError):
                unchanged_time = False
            if (
                _text(before_day) != _text(after_day)
                or not unchanged_time
                or _text(before_physio) != _text(after_name)
                or _text(specialty) not in {"ΦΘ", "ΡΟΜΠΟΤΙΚΗ"}
            ):
                blockers.append(f"R6 decision ledger row {number}: combined original was altered")
                continue
        matched.add(weekly_row)

    missing = len(controlled_weekly_rows - matched)
    if missing:
        blockers.append(f"{missing} R6 corrected/combined/held weekly rows lack exact ledger readback")



R6_ACTIVE_OVERLAPS_SHEET = "ΕΝΕΡΓΕΣ_ΕΠΙΚΑΛΥΨΕΙΣ"
R6_ACTIVE_OVERLAPS_HEADERS = (
    "Γραμμή", "Ασθενής", "Ημέρα", "Ώρα", "Ειδικότητες",
    "Κελιά", "Κατάσταση",
)
R6_COMBINED_DECISION = "ΕΓΚΕΚΡΙΜΕΝΟΣ ΣΥΝΔΥΑΣΜΟΣ — ΕΚΚΡΕΜΕΙ ΚΩΔΙΚΑΣ"
R6_PENDING_SPEECH_DECISION = "ΕΚΚΡΕΜΕΙ ΝΕΑ ΩΡΑ Λ/Θ"


def _audit_r6_active_overlaps(wb, proposals, blockers: list[str]) -> int | None:
    """Match the R6 active sheet to *all* held weekly collisions, 1:1.

    A valid approval status is not an admission or confirmed PatientID.
    The historical overlap list is never silently deleted or reinterpreted.
    """
    if R6_DECISIONS_SHEET not in wb.sheetnames:
        return None  # Original review format retains the old overlap policy.
    if R6_ACTIVE_OVERLAPS_SHEET not in wb.sheetnames:
        blockers.append("R6 active-overlap review sheet is missing")
        return None
    active = wb[R6_ACTIVE_OVERLAPS_SHEET]
    if tuple(active.cell(1, col).value for col in range(1, 8)) != R6_ACTIVE_OVERLAPS_HEADERS:
        blockers.append("R6 active-overlap headers differ from the reviewed schema")
        return None

    controlled = {
        "HOLD_APPROVED_COMBINED", "HOLD_CONFLICT", "HOLD_RESCHEDULE",
    }
    groups: dict[tuple, list[tuple]] = {}
    for number, values in proposals:
        if _text(values[10]) not in controlled:
            continue
        src_num = _source_row_number(values[1])
        try:
            clock = parse_clock(values[5])
        except (ProgrammeMappingError, TypeError, ValueError):
            clock = None
        if src_num is None or clock is None or not _text(values[2]):
            blockers.append(f"R6 held weekly row {number}: invalid overlap identity")
            continue
        key = (src_num, _text(values[2]), _text(values[4]), clock)
        groups.setdefault(key, []).append(values)

    seen: set[tuple] = set()
    unresolved = 0
    for number, values in _data_rows(active, 7, blockers=blockers):
        src_row, name, day, clock_raw, specialties, refs_raw, disposition = values
        src_num = _source_row_number(src_row)
        try:
            clock = parse_clock(clock_raw)
        except (ProgrammeMappingError, TypeError, ValueError):
            clock = None
        key = (src_num, _text(name), _text(day), clock)
        group = groups.get(key, ())
        exact_refs = _r6_exact_refs(refs_raw)
        expected_refs = []
        for proposal in group:
            expected_refs.extend(_r6_exact_refs(proposal[8]))
        expected_specialties = frozenset(_text(proposal[3]) for proposal in group)
        actual_specialties = frozenset(
            part.strip() for part in _text(specialties).split("+")
        )
        statuses = frozenset(_text(proposal[10]) for proposal in group)
        if (
            src_num is None or not _text(name) or day not in WEEKDAYS
            or clock is None or key in seen or len(group) != 2
            or not exact_refs or len(exact_refs) != len(set(exact_refs))
            or set(exact_refs) != set(expected_refs)
            or len(expected_refs) != len(set(expected_refs))
            or expected_specialties != actual_specialties
        ):
            blockers.append(f"R6 active-overlap row {number}: not an exact weekly group")
            continue
        seen.add(key)
        if statuses == {"HOLD_APPROVED_COMBINED"}:
            providers = {
                _provider_key(_text(proposal[6])) for proposal in group
            }
            if (
                _text(disposition) != R6_COMBINED_DECISION
                or expected_specialties != {"ΦΘ", "ΡΟΜΠΟΤΙΚΗ"}
                or len(providers) != 1 or "" in providers
            ):
                blockers.append(f"R6 active-overlap row {number}: combined pair differs")
        elif statuses == {"HOLD_CONFLICT", "HOLD_RESCHEDULE"}:
            if _text(disposition) != R6_PENDING_SPEECH_DECISION:
                blockers.append(f"R6 active-overlap row {number}: unresolved state differs")
            unresolved += 1
        else:
            blockers.append(f"R6 active-overlap row {number}: unsupported weekly hold states")

    if seen != set(groups):
        blockers.append(
            f"{len(set(groups) - seen)} R6 held weekly groups lack exact active-overlap review"
        )
    return unresolved



def _verify_source_bound_r6_statuses(
    *, programme: str | Path, wb, proposals, blockers: list[str],
) -> tuple[set[tuple], tuple[frozenset[int], ...]]:
    """Verify R6 corrected groups and exact combined pairs against original A:V.

    This is strictly an original-source proof, not PatientID/provider approval.
    Any absent/mutated source or ledger fails closed; all data stay local.
    """
    # Late import reuses the already-existing complete-group reconciliation
    # without adding a second source parser or a module-import cycle.
    from rehab_excel.trial_demo_source_reconciliation import (
        _source_bound_r6_corrected_groups,
        TrialDemoSourceReconciliationError,
    )

    original = load_workbook(programme, read_only=True, data_only=False)
    try:
        src = original[PROGRAMME_SHEET]
        try:
            groups = _source_bound_r6_corrected_groups(wb, src)
        except (TrialDemoSourceReconciliationError, KeyError, TypeError, ValueError):
            blockers.append("R6 corrected original-source groups require independent review")
            groups = {}
        corrected = {
            (source_row, treatment, day, clock, provider, source_cell)
            for (source_row, _), entries in groups.items()
            for treatment, day, clock, provider, source_cell in entries
        }
        active = wb[R6_ACTIVE_OVERLAPS_SHEET]
        weekly_groups: dict[tuple, list[tuple[int, tuple]]] = {}
        for number, values in proposals:
            if _text(values[10]) != "HOLD_APPROVED_COMBINED":
                continue
            source_row = _source_row_number(values[1])
            try:
                clock = parse_clock(values[5])
            except (ProgrammeMappingError, TypeError, ValueError):
                clock = None
            if source_row is None or clock is None:
                blockers.append(f"R6 combined weekly row {number}: invalid source slot")
                continue
            key = (source_row, _text(values[2]), _text(values[4]), clock)
            weekly_groups.setdefault(key, []).append((number, values))

        pairs: list[frozenset[int]] = []
        for number, values in _data_rows(active, 7, blockers=blockers):
            source_row, name, day, raw_clock, _, _, decision = values
            if _text(decision) != R6_COMBINED_DECISION:
                continue
            row_num = _source_row_number(source_row)
            try:
                clock = parse_clock(raw_clock)
            except (ProgrammeMappingError, TypeError, ValueError):
                clock = None
            key = (row_num, _text(name), _text(day), clock)
            members = weekly_groups.get(key, ())
            if row_num is None or clock is None or len(members) != 2:
                blockers.append(f"R6 combined active row {number}: source membership changed")
                continue
            if (
                {_text(member[3]) for _, member in members}
                != {"ΦΘ", "ΡΟΜΠΟΤΙΚΗ"}
                or any(_text(member[10]) != "HOLD_APPROVED_COMBINED"
                       for _, member in members)
            ):
                blockers.append(f"R6 combined active row {number}: invalid treatment pair")
                continue
            pair_valid = True
            for _, member in members:
                specialty = _text(member[3])
                treatment = SOURCE_TREATMENTS[specialty]
                clock_col, day_col, staff_col = (
                    ("C", "D", "E") if specialty == "ΦΘ"
                    else ("R", "S", "T")
                )
                original_cells = [
                    src.cell(row_num, ord(col) - 64)
                    for col in (clock_col, day_col, staff_col)
                ]
                if any(cell.data_type in {"f", "e"} for cell in original_cells):
                    pair_valid = False
                    break
                source_clock, source_days, source_provider = [
                    cell.value for cell in original_cells
                ]
                try:
                    source_candidates = make_weekly_entries(
                        patient_id="SOURCE-BOUND-CHECK",
                        treatment=treatment, source_cell=f"{clock_col}{row_num}",
                        clock=source_clock, weekdays=source_days,
                        therapist=source_provider,
                        implicit_daily=not _text(source_days),
                    )
                except (ProgrammeMappingError, TypeError, ValueError):
                    pair_valid = False
                    break
                if sum(
                    entry.weekday == _text(member[4])
                    and entry.start_time == clock
                    and _provider_key(entry.therapist_name)
                        == _provider_key(_text(member[6]))
                    for entry in source_candidates
                ) != 1:
                    pair_valid = False
                    break
            if not pair_valid:
                blockers.append(f"R6 combined active row {number}: original A:V slot differs")
                continue
            pairs.append(frozenset(member_number for member_number, _ in members))

        if (
            len(set(pairs)) != len(pairs)
            or set().union(*pairs) != {
                number for member_rows in weekly_groups.values()
                for number, _ in member_rows
            }
        ):
            blockers.append("R6 combined approvals have stale or incomplete source coverage")
        return corrected, tuple(pairs)
    finally:
        original.close()


def evaluate_trial_admission(
    path: str | Path, *, programme: str | Path | None = None,
) -> TrialAdmissionResult:
    """Perform full review, returning aggregate blockers without clinical names.

    The exact source row and hospital ward are used locally to join identities;
    names are cross-checked but never printed or included in the response.
    """
    # A complete original-source audit is mandatory before interpreting any
    # legacy shorthand. Without it, existing strict admission is unchanged.
    source_aligned = False
    if programme is not None:
        source_aligned = audit_trial_source_alignment(
            programme=programme, review=path,
        ).aligned
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        missing = REQUIRED_SHEETS.difference(wb.sheetnames)
        if missing:
            raise TrialAdmissionError("Missing review sheets: " + ", ".join(sorted(missing)))
        blockers: list[str] = []
        if programme is not None and not source_aligned:
            blockers.append("Original programme and review are not source-aligned")
        patients = tuple(_data_rows(wb[SOURCE_PATIENT_SHEET], 6, blockers=blockers))
        proposals = tuple(_data_rows(wb[SOURCE_WEEKLY_SHEET], 11, blockers=blockers))
        holds = tuple(_data_rows(wb[SOURCE_HOLDS_SHEET], 7, blockers=blockers))
        overlaps = tuple(_data_rows(wb[SOURCE_OVERLAP_SHEET], 7, blockers=blockers))
        _audit_r6_decision_readback(wb, proposals, blockers)
        r6_unresolved_overlaps = _audit_r6_active_overlaps(wb, proposals, blockers)
        corrected_verified: set[tuple] = set()
        exact_combined_rows: tuple[frozenset[int], ...] = ()
        if programme is not None and source_aligned and R6_DECISIONS_SHEET in wb.sheetnames:
            try:
                corrected_verified, exact_combined_rows = _verify_source_bound_r6_statuses(
                    programme=programme, wb=wb, proposals=proposals, blockers=blockers,
                )
            except (OSError, KeyError, ValueError, TypeError):
                blockers.append("R6 original source could not verify corrected/combined slots")

        patient_keys: dict[tuple[str, str, str], str] = {}
        verified_ids: list[str] = []
        seen_patient_rows: set[tuple[str, str]] = set()
        seen_original_rows: set[int] = set()
        seen_identities: set[tuple[str, str, str]] = set()
        for excel_row, row in patients:
            source_row, clinic, name, _ward, patient_id, confirmation = row
            key = (_text(clinic), _text(source_row), _text(name))
            location = key[:2]
            if not all(key):
                blockers.append(f"Patient staging row {excel_row}: missing source identity")
                continue
            source_number = _source_row_number(source_row)
            if source_number is None:
                blockers.append(f"Patient staging row {excel_row}: invalid original source row")
            elif source_number in seen_original_rows:
                blockers.append(f"Patient staging row {excel_row}: duplicate original source row")
            else:
                seen_original_rows.add(source_number)
            if location in seen_patient_rows:
                blockers.append(f"Patient staging row {excel_row}: duplicate clinic/source row")
            seen_patient_rows.add(location)
            if key in seen_identities:
                blockers.append(f"Patient staging row {excel_row}: repeated source identity")
            seen_identities.add(key)
            if not _text(patient_id) or _text(confirmation).upper() not in CONFIRMED_MARKERS:
                continue
            patient_keys[key] = _text(patient_id)
            verified_ids.append(_text(patient_id))

        duplicate_ids = len(verified_ids) - len({x.casefold() for x in verified_ids})
        if duplicate_ids:
            blockers.append(f"{duplicate_ids} duplicate confirmed PatientIDs")
        if len(verified_ids) != len(patients):
            blockers.append(
                f"{len(patients) - len(verified_ids)} patient IDs lack explicit confirmation"
            )
        if holds:
            blockers.append(f"{len(holds)} unresolved source cells")
        if R6_DECISIONS_SHEET in wb.sheetnames:
            if r6_unresolved_overlaps is not None and r6_unresolved_overlaps:
                blockers.append(f"{r6_unresolved_overlaps} unresolved active R6 overlap groups")
        elif overlaps:
            blockers.append(f"{len(overlaps)} cross-specialty overlap review rows")

        admitted: list[SourceWeeklyEntry] = []
        admitted_by_review_row: dict[int, SourceWeeklyEntry] = {}
        approved_combined_rows = set().union(*exact_combined_rows)
        unadmitted = 0
        for excel_row, row in proposals:
            clinic, source_row, name, specialty, day, clock, required_name, optional_name, cells, _raw, state = row
            key = (_text(clinic), _text(source_row), _text(name))
            patient_id = patient_keys.get(key)
            if not patient_id:
                unadmitted += 1
                continue
            status = _text(state)
            if status not in {
                "PROPOSED_EXACT_SOURCE", "PROPOSED_OPERATOR_CORRECTION",
                "HOLD_APPROVED_COMBINED",
            }:
                unadmitted += 1
                continue
            if status == "HOLD_APPROVED_COMBINED" and excel_row not in approved_combined_rows:
                unadmitted += 1
                continue
            treatment = SOURCE_TREATMENTS.get(_text(specialty))
            if treatment is None or _text(day) not in WEEKDAYS:
                blockers.append(f"Weekly staging row {excel_row}: unsupported specialty or day")
                unadmitted += 1
                continue
            original_number = _source_row_number(source_row)
            reference_text = _text(cells).upper()
            # The audited shorthand is unambiguous only for these exact
            # specialty column groups and the original, matching clock.
            # The independent source audit establishes all identity/clock
            # facts before we expand it to conventional full references.
            shorthand = {
                "ΕΡΓΟΘΕΡΑΠΕΙΑ": "G/H/I",
                "ΨΥΧΟΛΟΓΙΑ": "M/N/O",
            }
            if (
                programme is not None and source_aligned
                and reference_text == shorthand.get(_text(specialty))
                and original_number is not None
            ):
                reference_text = " / ".join(
                    f"{column}{original_number}"
                    for column in reference_text.split("/")
                )
            references = SRC_CELL_RE.findall(reference_text)
            allowed_columns = SOURCE_REF_COLUMNS[_text(specialty)]
            time_column = SOURCE_TIME_COLUMNS[_text(specialty)]
            if (
                original_number is None or not references
                or len(references) != len(set(references))
                or (time_column, str(original_number)) not in references
                or any(
                    column not in allowed_columns or int(number) != original_number
                    for column, number in references
                )
            ):
                blockers.append(f"Weekly staging row {excel_row}: invalid source cell references")
                unadmitted += 1
                continue
            provider = _text(required_name) if treatment in {"ΦΘ", "Ρομποτικό"} else _text(optional_name)
            if status == "PROPOSED_OPERATOR_CORRECTION" and (
                original_number, treatment, _text(day), parse_clock(clock)
                if _text(clock) else None, provider,
                f"{time_column}{original_number}",
            ) not in corrected_verified:
                blockers.append(f"Weekly staging row {excel_row}: correction not source-bound")
                unadmitted += 1
                continue
            if treatment in {"ΦΘ", "Ρομποτικό"} and not provider:
                blockers.append(f"Weekly staging row {excel_row}: mandatory provider missing")
                unadmitted += 1
                continue
            try:
                entry = SourceWeeklyEntry(
                    patient_id=patient_id,
                    treatment=treatment,
                    weekday=_text(day),
                    start_time=parse_clock(clock),
                    therapist_name=provider or None,
                    source_cell=f"{time_column}{original_number}",
                )
                admitted.append(entry)
                admitted_by_review_row[excel_row] = entry
            except (ProgrammeMappingError, ValueError):
                blockers.append(f"Weekly staging row {excel_row}: malformed weekday slot")
                unadmitted += 1

        if unadmitted:
            blockers.append(f"{unadmitted} weekly proposals cannot be admitted")
        approved_combined: frozenset[frozenset[SourceWeeklyEntry]] = frozenset()
        if exact_combined_rows:
            if any(not pair.issubset(admitted_by_review_row) for pair in exact_combined_rows):
                blockers.append("R6 combined pairs contain unadmitted reviewed rows")
            else:
                approved_combined = frozenset(
                    frozenset(admitted_by_review_row[index] for index in pair)
                    for pair in exact_combined_rows
                )
                if len(approved_combined) != len(exact_combined_rows):
                    blockers.append("R6 combined pair membership is duplicated")
        if admitted:
            # Explicit *exact membership* approval is required for shared
            # therapist slots. Approved patient IDs are never stored in GitHub,
            # and a partial/stale approval never admits additional patients.
            actual_groups: dict[tuple[str, str, object], set[str]] = {}
            names_by_key: dict[tuple[str, str, object], str] = {}
            for item in admitted:
                if item.therapist_name:
                    key = (_provider_key(item.therapist_name), item.weekday, item.start_time)
                    if not key[0]:
                        blockers.append("Provider label cannot be resolved by recurring conflict rules")
                        continue
                    actual_groups.setdefault(key, set()).add(item.patient_id)
                    names_by_key[key] = item.therapist_name

            approved: dict[tuple[str, str, object], frozenset[str]] = {}
            if SOURCE_SHARED_SHEET in wb.sheetnames:
                sheet = wb[SOURCE_SHARED_SHEET]
                headers = tuple(sheet.cell(1, col).value for col in range(1, 6))
                if headers != SHARED_HEADERS:
                    blockers.append("Shared-slot approval sheet headers are invalid")
                else:
                    for approval_row, values in _data_rows(sheet, 5, blockers=blockers):
                        provider, weekday, clock, ids, decision = values
                        try:
                            key = (_provider_key(_text(provider)), _text(weekday),
                                   parse_clock(clock))
                        except (ProgrammeMappingError, TypeError, ValueError):
                            blockers.append(
                                f"Shared-slot approval row {approval_row}: invalid time"
                            )
                            continue
                        member_ids = [_text(x) for x in _text(ids).split("|")]
                        if (
                            not key[0] or key[1] not in WEEKDAYS
                            or _text(decision) != SHARED_ACCEPTED
                            or len(member_ids) < 2
                            or any(not x for x in member_ids)
                            or len(set(member_ids)) != len(member_ids)
                            or key in approved
                        ):
                            blockers.append(
                                f"Shared-slot approval row {approval_row}: invalid or duplicated approval"
                            )
                            continue
                        approved[key] = frozenset(member_ids)
            collisions = {
                key: frozenset(members)
                for key, members in actual_groups.items() if len(members) > 1
            }
            unapproved = sum(
                1 for key, members in collisions.items()
                if approved.get(key) != members
            )
            stale = sum(
                1 for key, members in approved.items()
                if collisions.get(key) != members
            )
            if unapproved:
                blockers.append(f"{unapproved} unapproved shared provider weekday/time slots")
            if stale:
                blockers.append(f"{stale} stale or mismatched shared-slot approvals")
            try:
                validate_weekly_entry_consistency(
                    admitted, confirmed_combined_slots=approved_combined,
                )
            except ProgrammeMappingError:
                blockers.append("Patient cross-specialty duplicate slot detected")
        if not patients or not proposals:
            blockers.append("No source patients or weekly programme rows")
        return TrialAdmissionResult(
            candidate_patients=len(patients),
            verified_patient_ids=len(verified_ids),
            weekly_proposals=len(proposals),
            unadmitted_weekly_rows=unadmitted,
            unresolved_source_rows=len(holds),
            overlap_rows=(
                r6_unresolved_overlaps
                if R6_DECISIONS_SHEET in wb.sheetnames and r6_unresolved_overlaps is not None
                else len(overlaps)
            ),
            duplicate_ids=duplicate_ids,
            admitted_entries=tuple(admitted) if not blockers else (),
            blockers=tuple(blockers),
            confirmed_combined_slots=approved_combined if not blockers else frozenset(),
        )
    finally:
        wb.close()
