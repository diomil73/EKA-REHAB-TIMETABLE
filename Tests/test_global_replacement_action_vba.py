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


def test_global_resolved_plan_requires_explicit_operator_confirmation_before_persist():
    source = GLOBAL_REPLACEMENT_MODULE_CODE

    resolved_pos = source.index('If mode = "resolved" Then')
    confirm_pos = source.index("ConfirmGlobalReplacementProposal", resolved_pos)
    persist_pos = source.index("PersistReplacementPayload(payloadPath, True)", resolved_pos)

    assert resolved_pos < confirm_pos < persist_pos
    assert "If Not ConfirmGlobalReplacementProposal(responseText) Then GoTo CleanUp" in source
    assert 'If Not PersistReplacementPayload(payloadPath, mode = "resolved")' not in source


def test_global_replacement_confirmation_is_operator_facing_and_non_technical():
    source = GLOBAL_REPLACEMENT_MODULE_CODE
    messages = _messages(source)

    assert "Private Function ConfirmGlobalReplacementProposal" in source
    assert 'JsonLongValue(responseText, "normal_assignments")' in source
    assert "δεν έχει εφαρμοστεί ακόμη" in source
    assert "θέλετε να εφαρμοστεί" in source
    assert "backend" not in messages
    assert "payload" not in messages


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
