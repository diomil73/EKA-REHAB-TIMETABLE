import pytest

from rehab_excel.daily_change_ledger_in_place_vba import VBA_CODE as LEDGER_VBA_CODE
from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.operator_visibility_patch_vba import (
    patch_daily_input_patient_status_visibility,
    patch_daily_input_workload_visibility,
)
from tools.patch_native_restore_recovery import (
    NativeRestoreRecoveryPatchError,
    _compact_vba,
    _extract_vba_sub,
    patch_daily_change_ledger_code,
    patch_daily_input_status_rich_text,
)


def test_restore_uses_canonical_daily_input_patch_engine_and_declares_change_id():
    patched = patch_daily_change_ledger_code(LEDGER_VBA_CODE)
    assert "ApplyPayloadFileInPlace" not in patched
    assert "ApplyValidatedDailyInputPayload(patchPath)" in patched
    assert "Dim changeId As String" in patched
    assert 'Err.Raise vbObjectError + 2836, , "Restore patch apply failed"' in patched


def test_ledger_refresh_never_clears_arbitrary_block_below_column_a():
    patched = patch_daily_change_ledger_code(LEDGER_VBA_CODE)
    assert 'startRow = wsDaily.Cells(wsDaily.Rows.Count, 1).End(xlUp).Row + 3' not in patched
    assert 'Range("A" & startRow & ":G" & startRow + 200).ClearContents' not in patched
    assert 'wsDaily.Cells.Find(What:="*"' in patched
    assert "previousDisplayRow" in patched
    assert 'wsLedger.Cells(ledgerRow, 12).Value' in patched
    assert 'Range("A" & previousStartRow & ":G" & previousEndRow).ClearContents' in patched


def test_ledger_refresh_replacement_uses_procedure_boundaries_not_old_body_markers():
    start = LEDGER_VBA_CODE.index("Public Sub RefreshDailyChangeLedgerView()")
    end = LEDGER_VBA_CODE.index("End Sub", start) + len("End Sub")
    variant = (
        LEDGER_VBA_CODE[:start]
        + "Public Sub RefreshDailyChangeLedgerView()\r\n"
        + "    ' Excel COM may reflow or alter this entire body.\r\n"
        + "    MsgBox \"legacy placeholder\"\r\n"
        + "End Sub"
        + LEDGER_VBA_CODE[end:]
    )
    patched = patch_daily_change_ledger_code(variant)
    assert "legacy placeholder" not in patched
    assert 'wsDaily.Cells.Find(What:="*"' in patched
    assert "Public Sub RestoreDailyChangeRow()" in patched


def test_ledger_refresh_replacement_tolerates_crlf():
    variant = LEDGER_VBA_CODE.replace("\n", "\r\n")
    patched = patch_daily_change_ledger_code(variant)
    assert 'wsDaily.Cells.Find(What:="*"' in patched
    assert 'Range("A" & startRow & ":G" & startRow + 200).ClearContents' not in patched


def test_master_status_normalization_clears_character_level_strike_through():
    daily = patch_daily_input_workload_visibility(DAILY_INPUT_MODULE_CODE)
    daily = patch_daily_input_patient_status_visibility(daily)
    patched = patch_daily_input_status_rich_text(daily)
    assert "master.Cells(row, statusCol).Font.Strikethrough = False" in patched
    assert (
        "master.Cells(row, statusCol).Characters(1, "
        "Len(CStr(master.Cells(row, statusCol).Value))).Font.Strikethrough = False"
    ) in patched


def test_rich_text_verification_tolerates_com_spacing_and_line_continuation():
    variant = (
        "master.Cells(row, statusCol).Characters( 1 , _\r\n"
        "    Len( CStr( master.Cells(row, statusCol).Value ) ) ).Font.Strikethrough = False"
    )
    compact = _compact_vba(variant)
    assert ".Characters(" in compact
    assert ".Font.Strikethrough=False" in compact
    assert "statusCol" in compact


def test_master_status_verification_is_scoped_to_the_normalization_procedure():
    signature = "Private Sub NormalizeMasterPatientStatusPresentation(ByVal master As Worksheet)"
    code = (
        "Option Explicit\n"
        "Private Sub Unrelated()\n"
        "    Debug.Print \"Characters(.Font.Strikethrough=False)\"\n"
        "End Sub\n\n"
        + signature
        + "\n"
        "    master.Cells(2, 3).Characters(1, 4).Font.Strikethrough = False\n"
        "End Sub\n"
    )
    proc = _extract_vba_sub(code, signature)
    assert "Unrelated" not in proc
    compact = _compact_vba(proc)
    assert ".Characters(" in compact
    assert ".Font.Strikethrough=False" in compact


def test_recovery_patch_fails_closed_when_expected_markers_are_missing():
    with pytest.raises(NativeRestoreRecoveryPatchError):
        patch_daily_change_ledger_code("Option Explicit\n")
    with pytest.raises(NativeRestoreRecoveryPatchError):
        patch_daily_input_status_rich_text("Option Explicit\n")
