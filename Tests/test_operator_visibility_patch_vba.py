import inspect

from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.global_replacement_action_vba import GLOBAL_REPLACEMENT_MODULE_CODE
from rehab_excel.operator_visibility_patch_vba import (
    install_operator_visibility_patch,
    patch_daily_input_workload_visibility,
    patch_global_replacement_summary,
)
from tools import build_registration_menu_preview


def test_daily_input_refresh_stops_auto_rendering_workload_badges():
    patched = patch_daily_input_workload_visibility(DAILY_INPUT_MODULE_CODE)
    refresh_start = patched.index("Private Function RefreshOperationalViews() As Boolean")
    refresh_end = patched.index("Private Sub ValidatePatchTarget", refresh_start)
    refresh = patched[refresh_start:refresh_end]
    assert "ApplyMasterView" in refresh
    assert "RenderTherapistWorkloadStrip" not in refresh
    assert "therapist card" in refresh


def test_global_replacement_success_shows_new_assignment_details():
    patched = patch_global_replacement_summary(GLOBAL_REPLACEMENT_MODULE_CODE)
    assert "previousLastFactRow = LastReplacementFactRow()" in patched
    assert "BuildReplacementSummarySince(previousLastFactRow)" in patched
    assert "PatientDisplayName(patientId)" in patched
    assert 'replacementTherapist = Trim$(CStr(ws.Cells(row, 5).Value))' in patched
    assert 'replacementTime = Trim$(CStr(ws.Cells(row, 7).Text))' in patched
    assert '"Οι αναπληρώσεις εφαρμόστηκαν:"' in patched
    assert '" -> " & replacementTherapist & " -> " & replacementTime' in patched


def test_global_replacement_summary_only_reads_rows_appended_by_current_action():
    patched = patch_global_replacement_summary(GLOBAL_REPLACEMENT_MODULE_CODE)
    assert "For row = previousLastRow + 1 To lastRow" in patched
    assert 'statusText = UCase$(Trim$(CStr(ws.Cells(row, 10).Value)))' in patched
    assert 'If statusText = "ACTIVE" Then' in patched


def test_global_automatic_prompt_no_longer_claims_manual_routing():
    patched = patch_global_replacement_summary(GLOBAL_REPLACEMENT_MODULE_CODE)
    assert "Επιλέξτε Όχι για επιστροφή χωρίς εφαρμογή της πρότασης." in patched
    assert "Επιλέξτε Όχι αν θέλετε να κάνετε χειροκίνητες επιλογές" not in patched


def test_canonical_registration_build_installs_operator_visibility_patch():
    source = inspect.getsource(build_registration_menu_preview)
    assert "install_operator_visibility_patch" in source
    assert 'stage = "patching operator visibility"' in source
    assert source.index("install_therapist_workload_strip(vbproject, workbook)") < source.index(
        "install_operator_visibility_patch(vbproject)"
    )


def test_patch_installer_verifies_visibility_and_operator_contracts_semantically():
    source = inspect.getsource(install_operator_visibility_patch)
    assert "DAILY_INPUT still auto-renders workload badges" in source
    assert "DAILY_INPUT was not routed through the replacement operator" in source
    assert "Replacement summary patch was not installed" in source
    assert "Replacement summary procedure was not installed" in source
    assert '"PatientDisplayName(patientId)"' in source
    assert '"replacementTherapist"' in source
    assert '"replacementTime"' in source
    assert "install_replacement_operator(vbproject)" in source
    assert "→" not in source
