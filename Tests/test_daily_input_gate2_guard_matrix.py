"""Gate 2 fail-closed source contracts for exposed VBA mutation entrypoints.

These tests do not execute Excel/VBA. They are prerequisites, not substitutes,
for isolated Windows negative invocation tests.
"""
from __future__ import annotations

import re

import pytest

from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.daily_change_ledger_in_place_vba import VBA_CODE as LEDGER_CODE
from rehab_excel.global_replacement_action_vba import GLOBAL_REPLACEMENT_MODULE_CODE
from rehab_excel.double_booking_review_vba import MODULE_CODE as BOOKING_CODE
from rehab_excel.close_day_action_vba import CLOSE_DAY_MODULE_CODE
from rehab_excel.day_amendment_action_vba import DAY_AMENDMENT_MODULE_CODE
from rehab_excel.daily_input_journal_vba import JOURNAL_MODULE_CODE


def _procedure(source: str, name: str) -> str:
    names = re.findall(
        r"(?im)^\s*(?:Public|Private)\s+(?:Sub|Function)\s+([A-Za-z_][A-Za-z_0-9]*)\b",
        source,
    )
    assert names.count(name) == 1, f"Expected one entrypoint: {name}"
    match = re.search(
        rf"(?im)^\s*(?:Public|Private)\s+(Sub|Function)\s+{re.escape(name)}\b",
        source,
    )
    assert match is not None
    end = re.search(r"(?im)^\s*End\s+(?:Sub|Function)\s*$", source[match.end():])
    assert end is not None
    return source[match.start():match.end() + end.end()]


@pytest.mark.parametrize("source,name,guard_exit", [
    (DAILY_INPUT_MODULE_CODE, "ApplyDailyInputPreview", "Exit Sub"),
    (DAILY_INPUT_MODULE_CODE, "ApplyValidatedDailyInputPayload", "Exit Function"),
    (LEDGER_CODE, "PersistDailyChangePayload", "Err.Raise"),
    (LEDGER_CODE, "RunDailyChangeRestore", "Exit Function"),
    (LEDGER_CODE, "ApplyDailyChangeRestorePayload", "Exit Function"),
    (GLOBAL_REPLACEMENT_MODULE_CODE, "ResolveDailyReplacementsInPlace", "Exit Function"),
    (GLOBAL_REPLACEMENT_MODULE_CODE, "PersistReplacementPayload", "Exit Function"),
    (BOOKING_CODE, "ApplySelectedDoubleBooking", "Exit Function"),
    (BOOKING_CODE, "StartAuthoritativeCommit", "Exit Function"),
    (CLOSE_DAY_MODULE_CODE, "CloseOperationalDay", "Exit Sub"),
    (CLOSE_DAY_MODULE_CODE, "PersistCloseDayPayload", "Exit Function"),
    (DAY_AMENDMENT_MODULE_CODE, "StartClosedDayAmendment", "Exit Sub"),
    (DAY_AMENDMENT_MODULE_CODE, "PersistDaySavePayload", "Exit Function"),
], ids=lambda case: case if isinstance(case, str) and case in {"ApplyDailyInputPreview", "ApplyValidatedDailyInputPayload", "PersistDailyChangePayload", "RunDailyChangeRestore", "ApplyDailyChangeRestorePayload", "ResolveDailyReplacementsInPlace", "PersistReplacementPayload", "ApplySelectedDoubleBooking", "StartAuthoritativeCommit", "CloseOperationalDay", "PersistCloseDayPayload", "StartClosedDayAmendment", "PersistDaySavePayload"} else None)
def test_entrypoint_guards_present_before_first_work(source, name, guard_exit):
    body = _procedure(source, name)
    gate = re.search(r"(?i)If\s+Not\s+DailyInputJournalApplyEnabled\s*\(\s*\)\s+Then", body)
    assert gate is not None, f"Missing locked transaction gate in {name}"
    after_gate = body[gate.end():]
    # VBA supports both single-line If .. Then Exit Function and
    # multi-line If .. Then / Exit Function / End If.
    single_line = re.match(
        rf"[^\r\n]*\b{re.escape(guard_exit)}\b", after_gate
    )
    if single_line is None:
        end_if = re.search(r"(?im)^\s*End\s+If\s*$", after_gate)
        assert end_if is not None, f"Unterminated guard in {name}"
        assert re.search(
            rf"(?im)^\s*{re.escape(guard_exit)}\b",
            after_gate[:end_if.start()],
        ), f"Guard does not reject: {name}"
    # Must reject before disk operations, dialogs except its own warning,
    # subprocesses, workbooks/sheets writes, and calls to mutation helpers.
    prefix = body[:gate.start()]
    prefix = "\n".join(
        line for line in prefix.splitlines()[1:]
        if not line.lstrip().startswith("'")
    )
    assert not re.search(
        r"(?i)\b(?:ThisWorkbook\.Save|SaveCopyAs|\.Unprotect|"
        r"CreateObject\s*\(|Application\.Run|WScript\.Shell|"
        r"Persist(?:Replacement|CloseDay|DailyChange)|"
        r"ApplyValidatedDailyInputPayload|StartAuthoritativeCommit)\b",
        prefix,
    ), f"Side-effect preceding guard in {name}"


def test_journal_apply_gate_is_literal_false_without_extra_actions():
    function = _procedure(JOURNAL_MODULE_CODE, "DailyInputJournalApplyEnabled")
    statements = [
        x.strip() for x in function.splitlines()[1:-1]
        if x.strip() and not x.lstrip().startswith("'")
    ]
    assert statements == ["DailyInputJournalApplyEnabled = False"]


@pytest.mark.parametrize("name", [
    "RestoreDailyChangeRow",
    "RestoreSelectedDailyChanges",
    "RestoreAllDailyChanges",
])
def test_ledger_restore_ui_delegates_only_to_guarded_worker(name):
    body = _procedure(LEDGER_CODE, name)
    assert "RunDailyChangeRestore " in body
    assert _procedure(LEDGER_CODE, "RunDailyChangeRestore").find("DailyInputJournalApplyEnabled") >= 0


def test_booking_review_readonly_ui_is_distinct_from_write_action():
    review = _procedure(BOOKING_CODE, "ShowDoubleBookingReview")
    assert "ApplySelectedDoubleBooking" not in review
    assert "StartAuthoritativeCommit" not in review
