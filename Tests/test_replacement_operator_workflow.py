import inspect

from rehab_excel import replacement_operator_vba as operator_vba
from tools import prepare_replacement_operator_decision as decision_tool
from tools import validate_replacement_operator_choice as choice_tool


def test_operator_components_and_entrypoint_are_stable():
    assert operator_vba.MODULE_NAME == "modReplacementOperator"
    assert operator_vba.FORM_NAME == "frmReplacementDecision"
    assert operator_vba.MACRO_NAME == "ResolveReplacementWorkflow"


def test_router_offers_global_manual_or_cancel_before_resolution():
    source = operator_vba.MODULE_CODE
    assert "Public Function ResolveReplacementWorkflow() As Boolean" in source
    assert "vbYesNoCancel + vbQuestion" in source
    assert "ΝΑΙ: καλύτερη αυτόματη συνολική λύση" in source
    assert "ΟΧΙ: χειροκίνητη τακτοποίηση ασθενή-ασθενή" in source
    assert "ResolveDailyReplacementsInPlace()" in source
    assert "RunManualReplacementWorkflow()" in source


def test_manual_form_is_patient_by_patient_and_requires_confirmation():
    source = operator_vba.FORM_CODE
    assert "Επιλογή αναπλήρωσης" in source
    assert "Αρχικό πρόγραμμα:" in source
    assert "Επιβεβαίωση επιλογής" in source
    assert "Να αποθηκευτεί αυτή η επιλογή;" in source
    assert "ConfirmReplacementOperatorChoice" in source
    assert "Καλύτερη αυτόματη πρόταση" not in source


def test_manual_workflow_recomputes_after_every_confirmed_choice():
    source = operator_vba.MODULE_CODE
    assert "For iteration = 1 To 100" in source
    assert "PrepareNextReplacementDecision" in source
    assert "ApplyOperatorProjectionRefresh" in source
    assert "mDecisionAccepted = True" in source


def test_daily_input_route_patch_is_exact_and_fail_closed():
    source = inspect.getsource(operator_vba._patch_daily_input_route)
    assert 'legacy = "If ResolveDailyReplacementsInPlace() Then"' in source
    assert 'routed = "If ResolveReplacementWorkflow() Then"' in source
    assert "if legacy not in source" in source
    assert "raise RuntimeError" in source


def test_decision_preparer_is_read_only_and_uses_existing_global_solver_as_advisory():
    source = inspect.getsource(decision_tool.prepare_decision)
    assert "build_global_replacement_assignments" in source
    assert 'reason="operator proposal only"' in source
    assert "find_policy_replacement_options" in source
    assert "filter_replacement_options_for_service_hours" in source
    assert "file_sha256(source)" in source
    assert "append_daily_replacement_facts" not in source
    assert "save(" not in source


def test_decision_preparer_recalculates_from_active_replacements_and_prioritizes_scarce_sessions():
    source = inspect.getsource(decision_tool.prepare_decision)
    assert "read_daily_replacement_facts" in source
    assert "len(options_by_session[target.session_id])" in source
    assert "next_target = min(" in source


def test_choice_validator_reuses_existing_rules_without_writing_workbook():
    source = inspect.getsource(choice_tool.validate_choice)
    assert "find_policy_replacement_options" in source
    assert "filter_replacement_options_for_service_hours" in source
    assert "validate_policy_replacement_time" in source
    assert "create_replacement_assignment" in source
    assert "read_daily_replacement_facts" in source
    assert "file_sha256(source)" in source
    assert "append_daily_replacement_facts" not in source
    assert "save(" not in source


def test_operator_installer_patches_route_only_after_installing_form_and_module():
    source = inspect.getsource(operator_vba.install_replacement_operator)
    assert "module.CodeModule.AddFromString(MODULE_CODE)" in source
    assert "form.CodeModule.AddFromString(FORM_CODE)" in source
    assert "_patch_daily_input_route(vbproject)" in source
