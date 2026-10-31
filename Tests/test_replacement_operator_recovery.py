import inspect

from rehab_excel import replacement_operator_vba
from rehab_excel.operator_visibility_patch_vba import install_operator_visibility_patch
from tools import prepare_replacement_operator_decision, validate_replacement_operator_choice


def test_operator_restores_three_way_global_entry():
    code = replacement_operator_vba.MODULE_CODE
    assert "vbYesNoCancel" in code
    assert "καλύτερη αυτόματη συνολική λύση για όλους" in code
    assert "χειροκίνητη τακτοποίηση ασθενή-ασθενή" in code
    assert "ResolveDailyReplacementsInPlace()" in code
    assert "RunManualReplacementWorkflow()" in code


def test_manual_form_exposes_three_patient_outcomes_and_navigation():
    code = replacement_operator_vba.FORM_CODE
    assert 'cmdConfirm.Caption = "Αναπλήρωση"' in code
    assert 'cmdPostpone.Caption = "Αναβολή για σήμερα"' in code
    assert 'cmdNext.Caption = "Επόμενος"' in code
    assert 'cmdPrevious.Caption = "Προηγούμενος"' in code
    assert 'cmdAuto.Caption = "Αυτόματη τακτοποίηση υπολοίπων"' in code
    assert "PostponeReplacementOperatorTarget" in code
    assert "DeferReplacementOperatorTarget" in code
    assert "UndoPreviousReplacementOperatorDecision" in code


def test_manual_form_shows_therapist_absence_reason():
    assert '"absence_reason"' in replacement_operator_vba.FORM_CODE
    source = inspect.getsource(prepare_replacement_operator_decision.prepare_decision)
    assert '"absence_reason": _absence_reason' in source


def test_manual_choice_is_revalidated_and_auto_remainder_reuses_global_solver():
    code = replacement_operator_vba.MODULE_CODE
    assert 'ResolveProjectTool("validate_replacement_operator_choice.py")' in code
    assert "AutoResolveReplacementRemainder" in code
    assert "ResolveDailyReplacementsInPlace()" in code
    assert "mManualActions.Add" in code


def test_deferred_targets_are_not_marked_complete():
    source = inspect.getsource(prepare_replacement_operator_decision.prepare_decision)
    assert "skip_target_session_ids" in source
    assert "all_targets" in source
    assert "targets = tuple(target for target in all_targets if target.session_id not in skipped)" in source
    assert "if not targets:" in source
    assert "targets = all_targets" in source


def test_operator_refresh_stays_in_session_and_does_not_render_master_workload_ui():
    code = replacement_operator_vba.MODULE_CODE
    refresh_start = code.index("Private Function ApplyOperatorProjectionRefresh")
    refresh_end = code.index("Private Sub RollbackOperatorReplacement", refresh_start)
    refresh = code[refresh_start:refresh_end]
    assert "ApplyValidatedDailyInputPayload" in refresh
    assert "ApplyMasterView" in refresh
    assert "KeepApplicationShell" in refresh
    assert "RenderTherapistWorkloadStrip" not in refresh
    assert "ThisWorkbook.Close" not in code
    assert "Application.OnTime" not in code


def test_operator_installer_routes_daily_input_through_operator():
    source = inspect.getsource(replacement_operator_vba._patch_daily_input_route)
    assert 'legacy = "If ResolveDailyReplacementsInPlace() Then"' in source
    assert 'routed = "If ResolveReplacementWorkflow() Then"' in source
    patch_source = inspect.getsource(install_operator_visibility_patch)
    assert "install_replacement_operator(vbproject)" in patch_source


def test_operator_support_tools_are_read_only():
    prep = inspect.getsource(prepare_replacement_operator_decision)
    validate = inspect.getsource(validate_replacement_operator_choice)
    for source in (prep, validate):
        assert "Workbooks.Open" not in source
        assert "shutil.copy" not in source
        assert "save(" not in source
        assert "file_sha256(source)" in source
