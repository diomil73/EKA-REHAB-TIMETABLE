"""Synthetic active-census source formula and navigation safety checks."""
import inspect

import pytest

from rehab_excel import trial_demo_active_census as active
from rehab_excel import trial_demo_patient_exit_vba as exit_ui
from rehab_excel import trial_demo_native_seed as native
from rehab_excel.master_projection_writer import refresh_master_projection_in_place


def test_status_formulas_use_exact_patient_row_and_all_exit_rows():
    row = active.census_formulas(17, patient_type_source_column=9)
    assert row.identity == '=IF(PATIENTS!C17="","",PATIENTS!A17)'
    assert row.name == '=IF(PATIENTS!C17="","",PATIENTS!C17)'
    assert row.patient_type == '=IF(PATIENTS!C17="","",PATIENTS!I17)'
    assert row.room == '=IF(PATIENTS!C17="","",PATIENTS!B17)'
    assert "MATCH(A17&"",PATIENT_EXITS!$A:$A,0)" in row.last_active
    assert "INDEX(PATIENT_EXITS!$C:$C" in row.last_active
    assert 'IF(INT(E17)>=INT($H$2),"Ενεργός","Εξιτήριο / Λήξη")' in row.status


def test_empty_exit_means_active_and_last_active_date_is_inclusive():
    status = active.census_formulas(2, patient_type_source_column=6).status
    assert 'IF(E2="","Ενεργός"' in status
    assert 'INT(E2)>=INT($H$2)' in status
    assert status.count('"Ενεργός"') == 2


def test_dynamic_census_never_uses_fixed_exit_ledger_bound():
    formula = active.census_formulas(2, patient_type_source_column=6).last_active
    assert "$A:$A" in formula and "$C:$C" in formula
    assert "$2000" not in formula and "$1000" not in formula


@pytest.mark.parametrize("row,type_col", [(1, 6), (2, 5), (-1, 9)])
def test_unsafe_formula_grid_scope_fails_closed(row, type_col):
    with pytest.raises(active.DemoActiveCensusError):
        active.census_formulas(row, patient_type_source_column=type_col)


def test_census_opens_separately_with_master_back_navigation():
    assert active.SHEET_NAME == "DEMO_ACTIVE_CENSUS"
    assert active.MODULE_NAME != "modMasterToolbar"
    assert "Public Sub OpenDemoActiveCensus()" in active.VBA_CODE
    assert 'Worksheets("DEMO_ACTIVE_CENSUS")' in active.VBA_CODE
    assert ".Calculate" in active.VBA_CODE
    assert "Public Sub RefreshDemoActiveCensus()" in active.VBA_CODE
    assert "AutoFill" in active.VBA_CODE
    assert "patientLast > viewLast" in active.VBA_CODE
    assert "KeepApplicationShell" in active.VBA_CODE
    source = inspect.getsource(active.install_demo_active_census)
    assert 'caption="Ενεργοί ασθενείς"' in source
    assert 'action="GoToMaster"' in source
    assert 'ws.Cells(2, 8).Formula = "=TODAY()"' in source


def test_native_seed_and_exit_ui_use_same_census_without_rewriting_master():
    writer = inspect.getsource(native.Win32ComTrialDemoSeedBackend.write)
    seed = inspect.getsource(native.seed_trial_demo_workbook)
    code = exit_ui.VBA_CODE
    assert "install_demo_active_census(" in writer
    assert "verify_demo_active_census(" in seed
    assert "RefreshDemoActiveCensus" in code
    assert "MASTER_SCHEDULE" not in code
    assert "PATIENTS.Delete" not in code


def test_master_writer_supports_opt_in_date_effective_refresh_only():
    source = inspect.getsource(refresh_master_projection_in_place)
    assert "target_date: date | None = None" in source
    assert "build_master_projection(path, target_date=target_date)" in source
    assert "else build_master_projection(path)" in source
