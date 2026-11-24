import inspect

import pytest

from rehab_excel.daily_change_ledger_in_place_vba import (
    DailyChangeLedgerProvisionError,
    HEADERS,
    LEDGER_SHEET,
    MODULE_NAME,
    VBA_CODE,
    ensure_daily_change_ledger_sheet,
    install_daily_change_ledger,
    patch_daily_input_for_ledger,
)
from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.operator_visibility_patch_vba import (
    patch_daily_input_patient_status_visibility,
    patch_daily_input_rich_text_strike_contract,
    patch_daily_input_workload_visibility,
)
from tools import build_registration_menu_preview


def test_ledger_contract_has_stable_module_sheet_and_headers():
    assert MODULE_NAME == "modDailyChangeLedger"
    assert LEDGER_SHEET == "DAILY_CHANGE_LEDGER"
    assert HEADERS == (
        "ChangeID", "Date", "ChangeType", "SubjectID", "Subject", "FromText",
        "ToText", "Reason", "SourceRow", "Status", "RestoredAt", "DisplayRow",
    )


def test_ledger_ui_renders_change_history_and_all_restore_controls():
    for token in (
        '"Αλλαγές ημέρας"',
        '"Από"',
        '"Προς"',
        '"Αιτία"',
        '"Κατάσταση"',
        '"Επαναφορά"',
        '"Επαναφορά επιλεγμένων"',
        '"Επαναφορά όλων"',
        "Public Sub RestoreDailyChangeRow()",
        "Public Sub RestoreSelectedDailyChanges()",
        "Public Sub RestoreAllDailyChanges()",
    ):
        assert token in VBA_CODE


def test_ledger_view_never_blindly_clears_operational_rows():
    assert 'startRow = wsDaily.Cells(wsDaily.Rows.Count, 1).End(xlUp).Row + 3' not in VBA_CODE
    assert 'Range("A" & startRow & ":G" & startRow + 200).ClearContents' not in VBA_CODE
    assert 'Set lastDataCell = wsDaily.Cells.Find(What:="*"' in VBA_CODE
    assert 'If CStr(wsDaily.Cells(previousStartRow, 1).Value) = "Αλλαγές ημέρας" Then' in VBA_CODE
    assert 'wsDaily.Range("A" & previousStartRow & ":G" & previousEndRow).ClearContents' in VBA_CODE


def test_restore_is_hidden_in_session_and_uses_atomic_multi_id_preparer():
    assert 'WScript.Shell").Run(commandLine, 0, True)' in VBA_CODE
    assert "prepare_daily_change_restore_in_place.py" in VBA_CODE
    assert 'args = args & " --change-id " & QuoteArg(changeId)' in VBA_CODE
    assert "ApplyDailyChangeRestorePayload" in VBA_CODE
    assert "ThisWorkbook.Close" not in VBA_CODE
    assert "Application.OnTime" not in VBA_CODE
    assert "authoritative_commit_worker_cli.py" not in VBA_CODE


def test_restore_uses_canonical_validated_patch_engine_and_declares_change_id():
    assert "ApplyPayloadFileInPlace" not in VBA_CODE
    assert "ApplyValidatedDailyInputPayload(patchPath)" in VBA_CODE
    assert "Dim changeId As String" in VBA_CODE
    assert 'Err.Raise vbObjectError + 2836, , "Restore patch apply failed"' in VBA_CODE


def test_restore_payload_is_fail_closed_with_rollback_and_semantic_source_scope():
    assert 'PayloadValue(parts(1)) <> INPUT_SHEET' in VBA_CODE
    assert 'PayloadValue(parts(3)) <> "A"' in VBA_CODE
    assert 'PayloadValue(parts(4)) <> "E"' in VBA_CODE
    assert "ActiveLedgerChangeExists" in VBA_CODE
    assert "RestoreDailyChangeTransaction rollbackPath" in VBA_CODE
    assert 'RestoreOneSheetFromRollback rollbackPath, "THERAPIST_DAILY"' in VBA_CODE
    assert 'RestoreOneSheetFromRollback rollbackPath, "MASTER_SCHEDULE"' in VBA_CODE


def test_therapist_absence_restore_is_visible_but_not_falsely_enabled():
    assert 'changeType = "patient_absence" Or changeType = "session_cancellation"' in VBA_CODE
    assert '"Επαναφορά σύντομα"' in VBA_CODE


def test_daily_input_patch_accepts_change_rows_and_restores_ledger_on_failure():
    patched = patch_daily_input_for_ledger(DAILY_INPUT_MODULE_CODE)
    assert 'If CStr(parts(0)) = "CHANGE" Then' in patched
    assert "ValidateDailyChangePayload parts" in patched
    assert "PersistDailyChangePayload parts" in patched
    assert "RestoreDailyChangeLedgerFromRollback rollbackPath" in patched
    assert 'If CStr(parts(0)) <> "PATCH" Then GoTo RestoreNext' in patched
    assert "RefreshDailyChangeLedgerView" in patched


def test_ledger_patch_composes_after_operator_status_normalization():
    operator_patched = patch_daily_input_workload_visibility(DAILY_INPUT_MODULE_CODE)
    operator_patched = patch_daily_input_rich_text_strike_contract(operator_patched)
    operator_patched = patch_daily_input_patient_status_visibility(operator_patched)
    patched = patch_daily_input_for_ledger(operator_patched)
    expected = (
        "    ApplyMasterView\n"
        "    NormalizeMasterPatientStatusPresentation master\n"
        "    RefreshDailyChangeLedgerView\n"
        "    RefreshOperationalViews = True"
    )
    assert expected in patched
    assert patched.count("RefreshDailyChangeLedgerView") == 1
    assert "VerifyCharacterStrike(cell, True)" in patched


def test_canonical_build_installs_ledger_after_operator_patch():
    source = inspect.getsource(build_registration_menu_preview)
    assert "install_daily_change_ledger" in source
    assert 'stage = "installing daily change ledger"' in source
    assert source.index("install_operator_visibility_patch(vbproject)") < source.index(
        "install_daily_change_ledger(vbproject, workbook)"
    )


def test_installer_ensures_sheet_and_patches_daily_input():
    source = inspect.getsource(install_daily_change_ledger)
    assert "ensure_daily_change_ledger_sheet(workbook)" in source
    assert 'vbproject.VBComponents("modDailyInputAction")' in source
    assert "patch_daily_input_for_ledger(current)" in source


def test_restore_failure_keeps_snapshot_and_does_not_claim_rollback_is_verified():
    source = VBA_CODE.split(
        "Private Function ApplyDailyChangeRestorePayload", 1
    )[1].split("Private Function ActiveLedgerChangeExists", 1)[0]
    assert "Resume RollbackAttempt" in source
    assert "On Error GoTo RollbackFailed" in source
    assert "Resume RollbackUnverified" in source
    assert "Δεν έχει επαληθευτεί η επαναφορά" in source
    assert "Μην συνεχίσετε καταχωρίσεις" in source
    assert "οι αλλαγές αναιρέθηκαν" not in source
    assert "If ApplyDailyChangeRestorePayload Then" in source
    assert source.index("If ApplyDailyChangeRestorePayload Then") < source.index(
        "Kill rollbackPath"
    )
    assert "Kill rollbackPath" not in source.split("RollbackAttempt:", 1)[1].split(
        "CleanUp:", 1
    )[0]

def test_restore_and_ledger_writers_cannot_bypass_apply_lock():
    """All independently callable write paths must be fail-closed before mutation."""
    procedures = {
        "PersistDailyChangePayload": ("Public Sub PersistDailyChangePayload", "End Sub"),
        "RunDailyChangeRestore": ("Private Function RunDailyChangeRestore", "End Function"),
        "ApplyDailyChangeRestorePayload": ("Private Function ApplyDailyChangeRestorePayload", "End Function"),
        "RestoreDailyChangeLedgerFromRollback": ("Public Sub RestoreDailyChangeLedgerFromRollback", "End Sub"),
    }
    extracted = {}
    for name, (start, end) in procedures.items():
        assert VBA_CODE.count(start) == 1
        body = VBA_CODE.split(start, 1)[1].split(end, 1)[0]
        assert "DailyInputJournalApplyEnabled()" in body, name
        extracted[name] = body

    ledger = extracted["PersistDailyChangePayload"]
    assert ledger.index("DailyInputJournalApplyEnabled()") < ledger.index("ValidateDailyChangePayload parts")
    assert ledger.index("DailyInputJournalApplyEnabled()") < ledger.index("ws.Cells(row, 1).Value2")

    restore = extracted["RunDailyChangeRestore"]
    assert restore.index("DailyInputJournalApplyEnabled()") < restore.index("ThisWorkbook.Save")
    assert restore.index("DailyInputJournalApplyEnabled()") < restore.index("WScript.Shell")
    assert "Exit Function" in restore.split("If Not DailyInputJournalApplyEnabled() Then", 1)[1].split("End If", 1)[0]
    assert "Δεν έγιναν αλλαγές" in restore

    payload = extracted["ApplyDailyChangeRestorePayload"]
    assert payload.index("DailyInputJournalApplyEnabled()") < payload.index("ReadUtf8Text(payloadPath)")
    assert payload.index("DailyInputJournalApplyEnabled()") < payload.index("ThisWorkbook.SaveCopyAs")
    assert payload.index("DailyInputJournalApplyEnabled()") < payload.index("ClearContents")

    helper = extracted["RestoreDailyChangeLedgerFromRollback"]
    assert helper.index("DailyInputJournalApplyEnabled()") < helper.index("RestoreOneSheetFromRollback")
    assert 'Err.Raise vbObjectError + 2841' in helper


def test_restore_lock_uses_same_fixed_false_journal_gate_as_apply():
    from rehab_excel.daily_input_journal_vba import JOURNAL_MODULE_CODE
    source = JOURNAL_MODULE_CODE.split(
        "Public Function DailyInputJournalApplyEnabled()", 1
    )[1].split("End Function", 1)[0]
    assert "DailyInputJournalApplyEnabled = False" in source
    assert "DailyInputJournalApplyEnabled()" in DAILY_INPUT_MODULE_CODE
    assert "DailyInputJournalApplyEnabled()" in VBA_CODE

class _LedgerCell:
    def __init__(self, value=None):
        self.Value2 = value
        self.HasFormula = False


class _LedgerCells:
    def __init__(self):
        self.items = {}

    def __call__(self, row, col):
        return self.items.setdefault((row, col), _LedgerCell())

    def Clear(self):
        raise AssertionError("Installer must never clear the ledger")


class _LedgerSheet:
    def __init__(self, name, *, header=False):
        self.Name = name
        self.Visible = 2
        self.Cells = _LedgerCells()
        if header:
            for col, value in enumerate(HEADERS, 1):
                self.Cells(1, col).Value2 = value


class _LedgerSheets:
    def __init__(self, book):
        self.book = book
        self.fail_lookup = False
        self.add_calls = 0

    @property
    def Count(self):
        return len(self.book.sheets)

    def __call__(self, key):
        if self.fail_lookup:
            raise OSError("FICTIONAL DISCONNECTED COM OBJECT")
        if isinstance(key, int):
            return self.book.sheets[key - 1]
        return next(sheet for sheet in self.book.sheets if sheet.Name == key)

    def Add(self, *, After=None):
        self.add_calls += 1
        sheet = _LedgerSheet("SheetNew")
        self.book.sheets.append(sheet)
        return sheet


class _LedgerBook:
    def __init__(self, *, with_ledger=False):
        self.sheets = [_LedgerSheet("MASTER_SCHEDULE")]
        if with_ledger:
            self.sheets.append(_LedgerSheet(LEDGER_SHEET, header=True))
        self.Worksheets = _LedgerSheets(self)
        self.ReadOnly = False
        self.save_calls = 0

    def Save(self):
        self.save_calls += 1


def _ledger_history(book):
    ws = book.sheets[-1]
    ws.Cells(2, 1).Value2 = "FICTIONAL-CHANGE-ID"
    ws.Cells(2, 2).Value2 = "2026-11-24"
    ws.Cells(2, 3).Value2 = "patient_absence"
    ws.Cells(2, 4).Value2 = "FICTIONAL-P1"
    ws.Cells(2, 8).Value2 = "SYNTHETIC REASON"
    ws.Cells(2, 9).Value2 = 21
    ws.Cells(2, 10).Value2 = "RESTORED"
    ws.Cells(2, 11).Value2 = "2026-11-24T18:00:00"
    ws.Cells(2, 12).Value2 = 37
    return ws


def _sheet_values(sheet):
    return tuple(sorted(
        (row, col, cell.Value2, cell.HasFormula)
        for (row, col), cell in sheet.Cells.items.items()
    ))


def test_ledger_reinstall_never_resets_existing_restored_or_active_history():
    book = _LedgerBook(with_ledger=True)
    sheet = _ledger_history(book)
    # Extra presentation and historical row values are also retained exactly.
    sheet.Cells(3, 1).Value2 = "FICTIONAL-ACTIVE"
    sheet.Cells(3, 10).Value2 = "ACTIVE"
    before = _sheet_values(sheet)
    assert ensure_daily_change_ledger_sheet(book) is None
    assert ensure_daily_change_ledger_sheet(book) is None
    assert _sheet_values(sheet) == before
    assert sheet.Visible == 2
    assert book.Worksheets.add_calls == 0
    assert book.save_calls == 0


@pytest.mark.parametrize("kind", [
    "header", "header-formula", "visible",
])
def test_ledger_reinstall_rejects_bad_existing_contract_without_erasing_history(kind):
    book = _LedgerBook(with_ledger=True)
    sheet = _ledger_history(book)
    if kind == "header":
        sheet.Cells(1, 4).Value2 = "UNEXPECTED HEADER"
    elif kind == "header-formula":
        sheet.Cells(1, 4).HasFormula = True
    elif kind == "visible":
        sheet.Visible = -1
    before = _sheet_values(sheet)
    with pytest.raises(DailyChangeLedgerProvisionError, match="schema|veryHidden"):
        ensure_daily_change_ledger_sheet(book)
    assert _sheet_values(sheet) == before
    assert book.Worksheets.add_calls == 0
    assert book.save_calls == 0


def test_ledger_install_creates_only_missing_sheet_and_validates_readback():
    book = _LedgerBook()
    ensure_daily_change_ledger_sheet(book)
    assert book.Worksheets.add_calls == 1
    assert len(book.sheets) == 2
    sheet = book.sheets[-1]
    assert sheet.Name == LEDGER_SHEET
    assert sheet.Visible == 2
    assert tuple(sheet.Cells(1, col).Value2 for col in range(1, 13)) == HEADERS
    assert book.save_calls == 0
    before = _sheet_values(sheet)
    ensure_daily_change_ledger_sheet(book)
    assert _sheet_values(sheet) == before
    assert book.Worksheets.add_calls == 1


def test_ledger_install_does_not_treat_com_disconnect_as_missing_sheet():
    book = _LedgerBook(with_ledger=True)
    sheet = _ledger_history(book)
    before = _sheet_values(sheet)
    book.Worksheets.fail_lookup = True
    with pytest.raises(DailyChangeLedgerProvisionError, match="provisioning failed"):
        ensure_daily_change_ledger_sheet(book)
    assert _sheet_values(sheet) == before
    assert book.Worksheets.add_calls == 0
    assert book.save_calls == 0


def test_ledger_install_readonly_missing_sheet_fails_without_creating():
    book = _LedgerBook()
    book.ReadOnly = True
    with pytest.raises(DailyChangeLedgerProvisionError, match="read-only"):
        ensure_daily_change_ledger_sheet(book)
    assert len(book.sheets) == 1
    assert book.Worksheets.add_calls == 0
    assert book.save_calls == 0


class _LedgerCodeModule:
    def __init__(self, code):
        self.code = code

    @property
    def CountOfLines(self):
        return len(self.code.splitlines())

    def Lines(self, start, count):
        return self.code

    def DeleteLines(self, start, count):
        self.code = ""

    def AddFromString(self, value):
        self.code += value


class _LedgerComponent:
    def __init__(self, name, code):
        self.Name = name
        self.CodeModule = _LedgerCodeModule(code)


class _LedgerComponents:
    def __init__(self, initial):
        self.items = {name: _LedgerComponent(name, source) for name, source in initial.items()}
        self.add_calls = 0

    def __call__(self, name):
        # Excel VBComponent.Name changes after VBComponents.Add, and COM
        # resolves the new name immediately; this fake must do the same.
        for component in self.items.values():
            if component.Name == name:
                return component
        raise KeyError(name)

    def Add(self, kind):
        assert kind == 1
        self.add_calls += 1
        component = _LedgerComponent("NewComponent", "")
        self.items[component.Name] = component
        return component

    def Remove(self, component):
        self.items.pop(component.Name)


class _LedgerProject:
    def __init__(self, initial):
        self.VBComponents = _LedgerComponents(initial)


def test_ledger_vba_patch_preflight_occurs_before_any_sheet_or_module_change():
    book = _LedgerBook(with_ledger=True)
    sheet = _ledger_history(book)
    before = _sheet_values(sheet)
    project = _LedgerProject({"modDailyInputAction": "FICTIONAL NONPATCHABLE SOURCE"})
    with pytest.raises(RuntimeError, match="preflight marker"):
        install_daily_change_ledger(project, book)
    assert _sheet_values(sheet) == before
    assert project.VBComponents.add_calls == 0
    assert book.Worksheets.add_calls == 0
    assert book.save_calls == 0


def test_valid_ledger_install_preserves_old_history_and_updates_modules():
    book = _LedgerBook(with_ledger=True)
    sheet = _ledger_history(book)
    before = _sheet_values(sheet)
    project = _LedgerProject({"modDailyInputAction": DAILY_INPUT_MODULE_CODE})
    install_daily_change_ledger(project, book)
    assert _sheet_values(sheet) == before
    assert "PersistDailyChangePayload parts" in project.VBComponents(
        "modDailyInputAction"
    ).CodeModule.code
    assert "Public Sub RestoreDailyChangeRow()" in project.VBComponents(
        MODULE_NAME
    ).CodeModule.code
    assert project.VBComponents.add_calls == 1
    assert book.Worksheets.add_calls == 0
    assert book.save_calls == 0



def test_ledger_checkmark_is_generated_at_runtime_to_avoid_vbe_codepage_loss():
    assert 'wsDaily.Cells(startRow + 1, 1).Value = ChrW$(&H2713)' in VBA_CODE
    assert "✓" not in VBA_CODE
