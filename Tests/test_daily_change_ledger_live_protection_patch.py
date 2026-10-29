from rehab_excel.daily_change_ledger_live_protection_patch import patch_ledger_module_source
from rehab_excel.daily_change_ledger_vba import VBA_CODE


def test_patch_unprotects_before_render_and_reprotects_on_every_exit():
    patched = patch_ledger_module_source(VBA_CODE)

    assert "wsDaily.Unprotect" in patched
    assert "If startRow < 2 Then GoTo RenderProtectionExit" in patched
    assert "If lastLedgerRow < 2 Then GoTo RenderProtectionExit" in patched
    assert "RenderProtectionExit:" in patched
    assert "wsDaily.Protect UserInterfaceOnly:=True" in patched
    assert "If Not wsDaily Is Nothing Then wsDaily.Protect UserInterfaceOnly:=True" in patched


def test_patch_tolerates_extra_existing_declarations_in_refresh_procedure():
    variant = VBA_CODE.replace(
        "    On Error GoTo RenderError\n",
        "    Dim existingRepairState As Boolean\n    On Error GoTo RenderError\n",
        1,
    )

    patched = patch_ledger_module_source(variant)

    assert "Dim existingRepairState As Boolean" in patched
    assert "Dim renderErrorText As String" in patched
    assert "wsDaily.Unprotect" in patched
    assert "RenderProtectionExit:" in patched


def test_patch_is_idempotent():
    once = patch_ledger_module_source(VBA_CODE)
    twice = patch_ledger_module_source(once)
    assert twice == once
