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


def test_master_toolbar_contract_has_first_five_actions():
    assert MASTER_TOOLBAR_MODULE_NAME == "modMasterToolbar"
    assert MASTER_TOOLBAR_PREFIX == "ekaToolbar_"
    assert [item[1] for item in BUTTONS] == [
        "Νέος ασθενής",
        "Θεραπευτής / Φοιτητής",
        "Απουσία ασθενή",
        "Απουσία θεραπευτή",
        "Ημερήσιο πρόγραμμα",
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
