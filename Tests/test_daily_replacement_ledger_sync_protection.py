from rehab_excel.daily_replacement_ledger_sync_vba import MODULE_CODE


def test_runtime_ledger_refresh_temporarily_unprotects_therapist_daily():
    refresh_block = MODULE_CODE.split('If Not ApplyLedgerRows(rowsPath) Then GoTo SyncFailed', 1)[1]
    refresh_block = refresh_block.split('ThisWorkbook.Save', 1)[0]

    assert 'Set wsDaily = ThisWorkbook.Worksheets(DAILY_SHEET)' in refresh_block
    assert 'wsDaily.Unprotect' in refresh_block
    assert 'Application.Run "RefreshDailyChangeLedgerView"' in refresh_block
    assert 'wsDaily.Protect UserInterfaceOnly:=True' in refresh_block
    assert refresh_block.index('wsDaily.Unprotect') < refresh_block.index('Application.Run "RefreshDailyChangeLedgerView"')
    assert refresh_block.index('Application.Run "RefreshDailyChangeLedgerView"') < refresh_block.index('wsDaily.Protect UserInterfaceOnly:=True')


def test_runtime_ledger_sync_reprotects_daily_sheet_on_cleanup_or_error():
    sync_exit = MODULE_CODE.split('SyncExit:', 1)[1].split('End Function', 1)[0]

    assert 'On Error Resume Next' in sync_exit
    assert 'If Not wsDaily Is Nothing Then wsDaily.Protect UserInterfaceOnly:=True' in sync_exit
