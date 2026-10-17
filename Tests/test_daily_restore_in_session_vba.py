import pytest

from rehab_excel.daily_change_ledger_vba import VBA_CODE
from rehab_excel.daily_restore_in_session_vba import build_in_session_ledger_vba


def test_transformer_replaces_close_reopen_bridge_with_in_session_apply():
    transformed = build_in_session_ledger_vba(VBA_CODE)

    assert 'exitCode = CLng(shell.Run(commandLine, 0, True))' in transformed
    assert 'ApplyAuthoritativeRestorePreview previewPath' in transformed
    assert 'ThisWorkbook.SaveCopyAs rollbackPath' in transformed
    assert 'SyncRestoreSheetSet previewWorkbook, ThisWorkbook' in transformed
    assert 'ThisWorkbook.Close SaveChanges:=True' not in transformed
    assert 'Set exec = shell.Exec(commandLine)' not in transformed
    assert 'authoritative_commit_worker_cli.py' not in transformed


def test_in_session_restore_avoids_redundant_success_path_saves():
    transformed = build_in_session_ledger_vba(VBA_CODE)

    assert 'ApplyAuthoritativeRestorePreview previewPath\n    RefreshDailyChangeLedgerView\n    ThisWorkbook.Save' not in transformed
    assert 'ThisWorkbook.Save\n    ThisWorkbook.SaveCopyAs rollbackPath' not in transformed
    assert 'ThisWorkbook.SaveCopyAs rollbackPath' in transformed
    assert 'SyncRestoreSheetSet previewWorkbook, ThisWorkbook\n    ThisWorkbook.Save' in transformed


def test_in_session_sync_avoids_bloated_used_range_copy():
    transformed = build_in_session_ledger_vba(VBA_CODE)

    assert 'Set sourceRange = MeaningfulRestoreRange(sourceSheet)' in transformed
    assert 'sourceSheet.UsedRange' not in transformed
    assert 'Private Function MeaningfulRestoreRange(ByVal ws As Worksheet) As Range' in transformed
    assert 'SearchOrder:=xlByRows' in transformed
    assert 'SearchOrder:=xlByColumns' in transformed
    assert 'LookIn:=xlFormulas' in transformed


def test_transformer_keeps_atomic_selected_and_all_backend_arguments():
    transformed = build_in_session_ledger_vba(VBA_CODE)

    assert 'Call RunAuthoritativeRestore(replacementArgs, False)' in transformed
    assert 'Call RunAuthoritativeRestore("", True)' in transformed
    assert 'commandLine = commandLine & " --all-active"' in transformed
    assert 'replacementArgs = replacementArgs & " --replacement-id " & QuoteArg(replacementId)' in transformed


def test_success_message_is_user_facing_not_implementation_facing():
    transformed = build_in_session_ledger_vba(VBA_CODE)

    assert 'Η επαναφορά ολοκληρώθηκε επιτυχώς.' in transformed
    assert 'χωρίς κλείσιμο του αρχείου' not in transformed


def test_ledger_reason_uses_greek_business_label_for_legacy_technical_reasons():
    transformed = build_in_session_ledger_vba(VBA_CODE)

    assert (
        'wsDaily.Cells(displayRow, 5).Value = '
        'GreekLedgerReason(CStr(wsLedger.Cells(ledgerRow, 9).Value))'
    ) in transformed
    assert 'GreekLedgerReason = "Απουσία θεραπευτή"' in transformed
    assert 'Case "", "temporary replacement", "daily_input therapist absence", "daily input therapist absence"' in transformed
    assert 'wsDaily.Cells(displayRow, 5).Value = wsLedger.Cells(ledgerRow, 9).Value' not in transformed


def test_transformer_requires_known_bridge_markers():
    with pytest.raises(ValueError, match="markers were not found"):
        build_in_session_ledger_vba("Option Explicit")
