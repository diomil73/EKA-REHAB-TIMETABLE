import inspect

from rehab_excel.replacement_operator_vba import (
    FORM_CODE,
    FORM_NAME,
    MODULE_CODE,
    MODULE_NAME,
    _patch_daily_input_route,
)
from tools import prepare_replacement_operator_decision as decision_tool
from tools import validate_replacement_operator_choice as choice_tool


def test_operator_overlay_component_names_are_stable():
    assert MODULE_NAME == "modReplacementOperator"
    assert FORM_NAME == "frmReplacementDecision"


def test_operator_ui_requires_explicit_confirmation_and_supports_auto_proposal():
    assert 'cmdAuto.Caption = "Καλύτερη αυτόματη πρόταση"' in FORM_CODE
    assert 'cmdConfirm.Caption = "Επιβεβαίωση επιλογής"' in FORM_CODE
    assert "ConfirmReplacementOperatorChoice" in FORM_CODE
    assert "MsgBox(" in FORM_CODE
    assert "Να αποθηκευτεί αυτή η επιλογή;" in FORM_CODE


def test_operator_ui_does_not_auto_suspend_when_no_normal_choice_exists():
    assert "Δεν θα γίνει αυτόματη αναστολή ή διπλοκράτηση" in FORM_CODE
    assert "cmdConfirm.Enabled = False" in FORM_CODE
    assert "create_treatment_suspension" not in MODULE_CODE


def test_operator_workflow_recomputes_after_each_confirmed_choice():
    assert "For iteration = 1 To 100" in MODULE_CODE
    assert "PrepareNextReplacementDecision" in MODULE_CODE
    assert "ApplyOperatorProjectionRefresh" in MODULE_CODE
    assert "mDecisionAccepted = True" in MODULE_CODE


def test_daily_input_route_patch_is_exact_and_rejects_unknown_layouts():
    source = inspect.getsource(_patch_daily_input_route)
    assert 'legacy = "If ResolveDailyReplacementsInPlace() Then"' in source
    assert 'routed = "If RunReplacementOperatorWorkflow() Then"' in source
    assert "if legacy not in source" in source
    assert "raise RuntimeError" in source


def test_decision_preparer_uses_locked_global_solver_as_advisory_only():
    source = inspect.getsource(decision_tool.prepare_decision)
    assert "build_global_replacement_assignments" in source
    assert 'reason="operator proposal only"' in source
    assert "file_sha256(source)" in source
    assert "append_daily_replacement_facts" not in source
    assert "apply_write_plan" not in source


def test_decision_preparer_prioritizes_scarce_sessions_before_flexible_ones():
    source = inspect.getsource(decision_tool.prepare_decision)
    assert "len(options_by_session[target.session_id])" in source
    assert "next_target = min(" in source


def test_choice_validator_reuses_existing_policy_and_service_rules_without_writing():
    source = inspect.getsource(choice_tool.validate_choice)
    assert "find_policy_replacement_options" in source
    assert "filter_replacement_options_for_service_hours" in source
    assert "create_replacement_assignment" in source
    assert "read_daily_replacement_facts" in source
    assert "file_sha256(source)" in source
    assert "append_daily_replacement_facts" not in source
    assert "save(" not in source
