from rehab_excel.manual_replacement_operator_v2_patch import (
    V2_FORM_CODE,
    build_v2_module_code,
    patch_ledger_module_for_operator,
)
from rehab_excel.daily_change_ledger_vba import VBA_CODE as LEDGER_VBA_CODE


def test_v2_requires_explicit_provider_and_time_and_never_auto_advances_on_confirm():
    module = build_v2_module_code()
    form = V2_FORM_CODE

    assert "cmbProvider" in form
    assert "cmbTime" in form
    assert "cmbProvider.ListIndex = -1" in form
    assert "cmbTime.ListIndex = -1" in form
    assert "If modReplacementOperator.ConfirmReplacementOperatorChoice" in form
    assert "Then Unload Me" not in form.split("Private Sub cmdConfirm_Click()", 1)[1].split("End Sub", 1)[0]
    assert "RequestNextReplacementPatient" in module


def test_v2_exposes_treatment_suspension_as_first_class_manual_outcome():
    module = build_v2_module_code()
    form = V2_FORM_CODE

    assert "Αναβολή θεραπείας για σήμερα" in form
    assert "ConfirmOperatorTreatmentSuspension" in module
    assert 'ws.Cells(targetRow, 4).Value2 = "Αναστολή θεραπείας"' in module
    assert "ApplyOperatorProjectionRefresh()" in module


def test_v2_has_explicit_previous_next_change_and_automatic_remainder_controls():
    module = build_v2_module_code()
    form = V2_FORM_CODE

    assert "< Προηγούμενος ασθενής" in form
    assert "Επόμενος ασθενής >" in form
    assert "Αλλαγή τρέχουσας επιλογής" in form
    assert "Αυτόματη τακτοποίηση υπολοίπων" in form
    assert "RequestPreviousReplacementPatient" in module
    assert "RequestChangeCurrentOperatorDecision" in module
    assert "RequestAutomaticRemainingReplacements" in module
    assert "RunManualReplacementWorkflow = ResolveDailyReplacementsInPlace()" in module


def test_v2_keeps_confirmed_manual_facts_fixed_until_explicit_undo():
    module = build_v2_module_code()

    assert "mDecisionHistoryKinds.Add mCurrentDecisionKind" in module
    assert "mDecisionHistoryTokens.Add mCurrentDecisionToken" in module
    assert "UndoOperatorDecision" in module
    assert "RestoreReplacementForOperator" in module


def test_v2_projection_refresh_temporarily_unprotects_therapist_daily():
    module = build_v2_module_code()

    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY").Unprotect' in module
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY").Protect UserInterfaceOnly:=True' in module


def test_ledger_patch_exposes_authoritative_restore_wrapper_for_backward_navigation():
    patched = patch_ledger_module_for_operator(LEDGER_VBA_CODE)

    assert "Public Function RestoreReplacementForOperator" in patched
    assert 'RunAuthoritativeRestore(" --replacement-id " & QuoteArg(replacementId), False)' in patched
