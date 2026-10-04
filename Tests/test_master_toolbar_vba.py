import inspect

from rehab_excel.master_toolbar_vba import (
    BUTTONS,
    MASTER_TOOLBAR_MODULE_CODE,
    MASTER_TOOLBAR_MODULE_NAME,
    MASTER_TOOLBAR_PREFIX,
    _ensure_app_header_rows,
)
from rehab_excel.registration_menu_vba import Win32ComRegistrationMenuInstaller


class _Cell:
    def __init__(self, value=None):
        self.Value = value


class _Rows:
    def __init__(self, ws, key):
        self.ws = ws
        self.key = key

    @property
    def RowHeight(self):
        return self.ws.row_heights.get(self.key)

    @RowHeight.setter
    def RowHeight(self, value):
        self.ws.row_heights[self.key] = value

    def Insert(self):
        self.ws.shifted = True
        shifted = {}
        for (row, col), value in self.ws.values.items():
            shifted[(row + 3, col)] = value
        self.ws.values = shifted


class _FakeWorksheet:
    def __init__(self, header_row=1):
        self.values = {
            (header_row, 2): "Θάλαμος",
            (header_row, 3): "Ασθενής",
        }
        self.row_heights = {}
        self.shifted = False

    def Cells(self, row, col):
        return _Cell(self.values.get((row, col)))

    def Rows(self, key):
        return _Rows(self, key)


def test_master_toolbar_contract_has_primary_actions_and_exit():
    assert MASTER_TOOLBAR_MODULE_NAME == "modMasterToolbar"
    assert MASTER_TOOLBAR_PREFIX == "ekaToolbar_"
    assert [item[1] for item in BUTTONS] == [
        "Νέος ασθενής",
        "Θεραπευτής / Φοιτητής",
        "Απουσία ασθενή",
        "Απουσία θεραπευτή",
        "Ημερήσιο πρόγραμμα",
        "Save & Exit",
    ]


def test_master_toolbar_actions_route_to_existing_features():
    assert "frmNewPatient.Show" in MASTER_TOOLBAR_MODULE_CODE
    assert "ShowRegistrationMenu" in MASTER_TOOLBAR_MODULE_CODE
    assert 'GoToDailyInputSection "ΑΠΟΥΣΙΕΣ / ΑΚΥΡΩΣΕΙΣ ΑΣΘΕΝΩΝ"' in MASTER_TOOLBAR_MODULE_CODE
    assert 'GoToDailyInputSection "ΑΠΟΥΣΙΕΣ ΘΕΡΑΠΕΥΤΩΝ"' in MASTER_TOOLBAR_MODULE_CODE
    assert 'Worksheets("THERAPIST_DAILY").Activate' in MASTER_TOOLBAR_MODULE_CODE


def test_app_header_rows_are_inserted_above_legacy_master_header():
    ws = _FakeWorksheet(header_row=1)

    header_row = _ensure_app_header_rows(ws)

    assert header_row == 4
    assert ws.shifted is True
    assert ws.values[(4, 2)] == "Θάλαμος"
    assert ws.values[(4, 3)] == "Ασθενής"
    assert ws.row_heights[1] == 6
    assert ws.row_heights[2] == 34
    assert ws.row_heights[3] == 6


def test_existing_app_header_is_idempotent():
    ws = _FakeWorksheet(header_row=4)

    header_row = _ensure_app_header_rows(ws)

    assert header_row == 4
    assert ws.shifted is False


def test_registration_menu_installer_installs_master_toolbar():
    source = inspect.getsource(Win32ComRegistrationMenuInstaller.install)
    assert "install_master_toolbar(vbproject, workbook)" in source
    assert "MASTER toolbar module was not confirmed" in source



def test_master_toolbar_includes_exit_action():
    captions = [item[1] for item in BUTTONS]
    assert "Save & Exit" in captions
    assert "Public Sub ExitApplication()" in MASTER_TOOLBAR_MODULE_CODE
    assert "RestoreExcelInterface" in MASTER_TOOLBAR_MODULE_CODE
    assert "ThisWorkbook.Close SaveChanges:=True" in MASTER_TOOLBAR_MODULE_CODE


def test_master_toolbar_includes_return_to_master_action():
    assert "Public Sub GoToMaster()" in MASTER_TOOLBAR_MODULE_CODE
    assert 'Worksheets("MASTER_SCHEDULE").Activate' in MASTER_TOOLBAR_MODULE_CODE
    assert "KeepApplicationShell" in MASTER_TOOLBAR_MODULE_CODE


def test_master_toolbar_freezes_through_patient_name_column():
    import inspect
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._freeze_master_identity_columns)
    assert "window.SplitColumn = 3" in source
    assert "window.FreezePanes = True" in source


def test_operational_sheets_receive_master_and_exit_navigation():
    import inspect
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._install_operational_navigation)
    assert '("DAILY_INPUT", "THERAPIST_DAILY")' in source
    assert 'caption="← MASTER"' in source
    assert 'macro="GoToMaster"' in source
    assert 'caption="Save & Exit"' in source
    assert 'macro="ExitApplication"' in source



def test_operational_navigation_is_placed_below_used_rows():
    import inspect
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._install_operational_navigation)
    assert "last_row = _last_used_row(ws)" in source
    assert "anchor_row = last_row + 2" in source
    assert 'ws.Range(f"A{anchor_row}")' in source


def test_user_facing_navigation_scope_excludes_master():
    from rehab_excel.master_toolbar_vba import USER_FACING_SHEETS

    assert USER_FACING_SHEETS == ("DAILY_INPUT", "THERAPIST_DAILY")
    assert "MASTER_SCHEDULE" not in USER_FACING_SHEETS


def test_navigation_finalizer_reapplies_after_sheet_creation():
    import inspect
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba.finalize_user_navigation)
    assert "_install_operational_navigation(workbook)" in source
    assert "workbook.Save()" in source
