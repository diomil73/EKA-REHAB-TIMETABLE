from rehab_excel.global_replacement_action_vba import (
    GLOBAL_REPLACEMENT_MACRO_NAME,
    GLOBAL_REPLACEMENT_MODULE_CODE,
    GLOBAL_REPLACEMENT_MODULE_NAME,
)


def _messages(source: str) -> str:
    return "\n".join(line for line in source.splitlines() if "MsgBox" in line).casefold()


def test_global_replacement_action_names_are_stable():
    assert GLOBAL_REPLACEMENT_MODULE_NAME == "modGlobalReplacementAction"
    assert GLOBAL_REPLACEMENT_MACRO_NAME == "ResolveDailyReplacementsInPlace"


def test_global_replacement_action_is_in_place_and_hidden():
    source = GLOBAL_REPLACEMENT_MODULE_CODE

    assert "prepare_global_replacements_in_place.py" in source
    assert 'CreateObject("WScript.Shell").Run(commandLine, 0, True)' in source
    assert "PersistReplacementPayload" in source
    assert "PersistReplacementFacts" in source
    assert "PersistSuspensions" in source
    assert "ApplyValidatedDailyInputPayload" in source
    assert "ThisWorkbook.Close" not in source
    assert "Application.OnTime" not in source
    assert "authoritative_commit_worker_cli.py" not in source


def test_global_replacement_action_uses_guided_review_not_raw_replacements_sheet():
    source = GLOBAL_REPLACEMENT_MODULE_CODE

    assert 'mode = "review_required"' in source
    assert 'Application.Run "ShowDoubleBookingReview"' in source
    assert 'Worksheets("REPLACEMENTS").Activate' not in source


def test_global_replacement_action_requires_operator_confirmation_before_persisting():
    source = GLOBAL_REPLACEMENT_MODULE_CODE

    gate = "If Not ConfirmAutomaticReplacementProposal(mode) Then GoTo CleanUp"
    persist = "If Not PersistReplacementPayload(payloadPath, mode = \"resolved\") Then GoTo CleanUp"

    assert "Private Function ConfirmAutomaticReplacementProposal" in source
    assert gate in source
    assert persist in source
    assert source.index(gate) < source.index(persist)
    assert 'vbYesNo + vbQuestion' in source
    assert '"Εφαρμογή καλύτερης αυτόματης λύσης"' in source


def test_global_replacement_action_has_transaction_rollback_and_diagnostic():
    source = GLOBAL_REPLACEMENT_MODULE_CODE
    messages = _messages(source)

    assert "ThisWorkbook.SaveCopyAs rollbackPath" in source
    assert "RestoreReplacementTransaction" in source
    assert "eka_replacement_last_error.log" in source
    assert 'Print #handle, "phase=" & phase' in source
    assert "eka_replacement_last_error.log" not in messages


def test_global_replacement_operator_messages_hide_internal_details():
    messages = _messages(GLOBAL_REPLACEMENT_MODULE_CODE)

    assert "backend" not in messages
    assert "worker" not in messages
    assert "preview" not in messages
    assert "hash" not in messages


def test_replacement_transaction_preserves_recovery_copy_and_fails_closed():
    source = GLOBAL_REPLACEMENT_MODULE_CODE
    block = source.split("Private Function PersistReplacementPayload", 1)[1].split(
        "Private Function PersistReplacementFacts", 1
    )[0]
    assert "Resume RecoveryAttempt" in block
    assert "On Error GoTo RecoveryFailed" in block
    assert "Resume RecoveryUnverified" in block
    assert "Err.Raise vbObjectError + 2513" in block
    assert "Δεν έχει επαληθευτεί πλήρης επαναφορά" in block
    assert "Μην συνεχίσετε καταχωρίσεις" in block
    assert "οι αλλαγές αναιρέθηκαν" not in block
    assert "previousEvents = Application.EnableEvents" in block.split(
        'phase = "read payload"', 1
    )[0]
    assert "previousScreenUpdating = Application.ScreenUpdating" in block.split(
        'phase = "read payload"', 1
    )[0]
    cleanup = block.split("PersistExit:", 1)[1]
    assert "If PersistReplacementPayload Then" in cleanup
    assert cleanup.index("If PersistReplacementPayload Then") < cleanup.index("Kill rollbackPath")
    assert "Kill rollbackPath" not in block.split("RecoveryAttempt:", 1)[1].split(
        "PersistExit:", 1
    )[0]
    assert "PersistReplacementPayload = True" in block
