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
    assert "ThisWorkbook.Save" in MASTER_TOOLBAR_MODULE_CODE
    assert "Application.EnableEvents = False" in MASTER_TOOLBAR_MODULE_CODE
    assert "ThisWorkbook.Close SaveChanges:=False" in MASTER_TOOLBAR_MODULE_CODE


def test_master_toolbar_includes_return_to_master_action():
    assert "Public Sub GoToMaster()" in MASTER_TOOLBAR_MODULE_CODE
    assert 'Worksheets("MASTER_SCHEDULE").Activate' in MASTER_TOOLBAR_MODULE_CODE
    assert "KeepApplicationShell" in MASTER_TOOLBAR_MODULE_CODE


def test_master_toolbar_clears_freeze_and_split():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._clear_master_freeze)
    assert "window.FreezePanes = False" in source
    assert "window.SplitColumn = 0" in source
    assert "window.SplitRow = 0" in source


def test_master_toolbar_forces_black_button_text():
    from rehab_excel import master_toolbar_vba

    helper = inspect.getsource(master_toolbar_vba._set_shape_text_black)
    install = inspect.getsource(master_toolbar_vba.install_master_toolbar)
    nav = inspect.getsource(master_toolbar_vba._add_navigation_button)
    assert "ForeColor.RGB = 0" in helper
    assert "_set_shape_text_black(shape)" in install
    assert "_set_shape_text_black(shape)" in nav


def test_operational_sheets_receive_master_and_exit_navigation():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._install_operational_navigation)
    assert "for sheet_name in USER_FACING_SHEETS" in source
    assert 'caption="← MASTER"' in source
    assert 'macro="GoToMaster"' in source
    assert 'caption="Save & Exit"' in source
    assert 'macro="ExitApplication"' in source


def test_operational_navigation_is_placed_below_used_rows():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._install_operational_navigation)
    assert "last_row = _last_used_row(ws)" in source
    assert "anchor_row = last_row + 2" in source
    assert 'ws.Range(f"A{anchor_row}")' in source


def test_user_facing_navigation_scope_excludes_master():
    from rehab_excel.master_toolbar_vba import USER_FACING_SHEETS

    assert USER_FACING_SHEETS == ("DAILY_INPUT", "THERAPIST_DAILY", "REPLACEMENTS")
    assert "MASTER_SCHEDULE" not in USER_FACING_SHEETS


def test_navigation_finalizer_reapplies_after_sheet_creation():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba.finalize_user_navigation)
    assert "_install_operational_navigation(workbook)" in source
    assert "workbook.Save()" in source


def test_daily_input_footer_has_apply_and_continue_action():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._install_operational_navigation)
    assert 'sheet_name == "DAILY_INPUT"' in source
    assert 'caption="Εφαρμογή & Συνέχεια"' in source
    assert 'macro="ApplyDailyInputPreview"' in source


def test_replacements_sheet_receives_footer_navigation():
    from rehab_excel.master_toolbar_vba import USER_FACING_SHEETS
    assert "REPLACEMENTS" in USER_FACING_SHEETS


def test_master_display_schema_adds_psychology_and_afternoon_columns():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._ensure_master_display_schema)
    assert 'ws.Columns("K:L").Insert()' in source
    assert '"ΨΥΧΟΛΟΓΟΙ"' in source
    assert '"ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ"' in source
    assert '"Κατάσταση"' in source


def test_master_display_uses_compact_recliner_column():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._ensure_master_display_schema)
    assert '"Ανακλ/μενο"' in source
    assert 'ws.Columns("G").ColumnWidth = 11' in source
    assert 'ws.Columns("G").WrapText = True' in source


def test_master_display_clears_false_only_on_room_separator_rows():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._ensure_master_display_schema)
    assert "if room and not patient" in source
    assert 'casefold() == "false"' in source
    assert "for col in range(13, 17)" in source
    assert "ws.Cells(row, col).ClearContents()" in source


def test_master_new_specialties_have_distinct_light_colors():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._ensure_master_display_schema)
    assert 'ws.Range(f"K{header_row}:K{last_row}").Interior.Color = 13421823' in source
    assert 'ws.Range(f"L{header_row}:L{last_row}").Interior.Color = 16764108' in source


def test_save_exit_saves_before_restoring_and_closing():
    save_pos = MASTER_TOOLBAR_MODULE_CODE.index("ThisWorkbook.Save")
    restore_pos = MASTER_TOOLBAR_MODULE_CODE.index("RestoreExcelInterface")
    close_pos = MASTER_TOOLBAR_MODULE_CODE.index("ThisWorkbook.Close SaveChanges:=False")
    assert save_pos < restore_pos < close_pos
    assert 'MsgBox "Δεν ήταν δυνατή η αποθήκευση και έξοδος:' in MASTER_TOOLBAR_MODULE_CODE


def test_master_scroll_area_stops_near_last_patient_row():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._ensure_master_display_schema)
    assert 'ws.ScrollArea = f"A1:M{last_row + 1}"' in source
    assert 'ws.Rows(f"{last_row + 2}:{int(ws.Rows.Count)}").Hidden = True' in source


def test_master_specialty_colors_do_not_fill_entire_columns():
    from rehab_excel import master_toolbar_vba

    source = inspect.getsource(master_toolbar_vba._ensure_master_display_schema)
    assert 'ws.Columns("K").Interior.Color' not in source
    assert 'ws.Columns("L").Interior.Color' not in source


def test_save_exit_temporarily_disables_workbook_events():
    assert "eventsWereEnabled = Application.EnableEvents" in MASTER_TOOLBAR_MODULE_CODE
    assert "Application.EnableEvents = False" in MASTER_TOOLBAR_MODULE_CODE
    assert "Application.EnableEvents = eventsWereEnabled" in MASTER_TOOLBAR_MODULE_CODE


def test_master_vba_view_hides_rows_below_last_written_plus_one():
    assert "Public Sub ApplyMasterView()" in MASTER_TOOLBAR_MODULE_CODE
    assert "lastRoomRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row" in MASTER_TOOLBAR_MODULE_CODE
    assert "lastPatientRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row" in MASTER_TOOLBAR_MODULE_CODE
    assert "firstHiddenRow = lastRow + 2" in MASTER_TOOLBAR_MODULE_CODE
    assert 'ws.ScrollArea = "A1:M" & CStr(lastRow + 1)' in MASTER_TOOLBAR_MODULE_CODE
    assert 'ws.Rows(CStr(firstHiddenRow) & ":" & CStr(ws.Rows.Count)).Hidden = True' in MASTER_TOOLBAR_MODULE_CODE


def test_save_exit_quits_excel_when_app_workbook_is_only_workbook():
    assert "workbookCount = Application.Workbooks.Count" in MASTER_TOOLBAR_MODULE_CODE
    assert "If workbookCount <= 1 Then" in MASTER_TOOLBAR_MODULE_CODE
    assert "Application.Quit" in MASTER_TOOLBAR_MODULE_CODE
    assert "Else" in MASTER_TOOLBAR_MODULE_CODE
    assert "ThisWorkbook.Close SaveChanges:=False" in MASTER_TOOLBAR_MODULE_CODE
