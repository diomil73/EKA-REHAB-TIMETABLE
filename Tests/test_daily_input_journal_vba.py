from rehab_excel.application_contract import REQUIRED_APP_COMPONENTS
from rehab_excel.daily_input_journal_vba import (
    JOURNAL_MODULE_CODE,
    JOURNAL_MODULE_NAME,
    install_daily_input_journal,
)


def test_prepared_bridge_is_a_distinct_required_vba_component():
    assert JOURNAL_MODULE_NAME == "modDailyInputJournal"
    assert JOURNAL_MODULE_NAME in REQUIRED_APP_COMPONENTS


def test_vba_bridge_validates_row_and_protected_journal_before_append():
    source = JOURNAL_MODULE_CODE
    for marker in (
        'ThisWorkbook.ReadOnly',
        'CStr(fields(1)) <> "0"',
        'CStr(fields(2)) <> "PREPARED"',
        'IsJournalDigest(CStr(fields(3)))',
        'Left$(CStr(fields(4)), Len(JOURNAL_PREFIX)) <> JOURNAL_PREFIX',
        'CLng(ws.Visible) <> XL_SHEET_VERY_HIDDEN',
        'Not CBool(ws.ProtectContents)',
        'ws.Unprotect',
        'ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True',
        'ThisWorkbook.Save',
        '--verify-handoff',
        '--native-excel',
    ):
        assert marker in source


def test_vba_bridge_never_deletes_or_rewrites_a_persisted_journal_row():
    source = JOURNAL_MODULE_CODE
    assert ".ClearContents" not in source
    assert "ws.Rows(" not in source
    assert "ws.Cells(targetRow" in source
    assert "Never delete or rewrite a possibly durable journal row" in source
    assert "PersistPreparedDailyInputJournal = False" in source


def test_probe_only_persists_prepared_and_never_calls_operational_apply():
    source = JOURNAL_MODULE_CODE
    probe = source.split("Public Sub ProbeDailyInputPreparedJournal", 1)[1].split(
        "End Sub", 1
    )[0]
    assert "PersistPreparedDailyInputJournal" in probe
    assert "ApplyValidatedDailyInputPayload" not in probe
    assert "No patient or treatment changes were applied." in probe
    assert "RunPreparedDailyInputJournalProbe" in source


def test_apply_gate_is_explicitly_disabled_until_terminal_lifecycle_is_implemented():
    source = JOURNAL_MODULE_CODE
    gate = source.split("Public Function DailyInputJournalApplyEnabled()", 1)[1].split(
        "End Function", 1
    )[0]
    assert "DailyInputJournalApplyEnabled = False" in gate


def test_installer_adds_named_module_and_full_source():
    added = []

    class FakeCodeModule:
        def AddFromString(self, source):
            self.source = source

    class FakeComponent:
        def __init__(self):
            self.CodeModule = FakeCodeModule()
            self.Name = ""

    class FakeComponents:
        def __call__(self, name):
            raise KeyError(name)

        def Add(self, component_type):
            assert component_type == 1
            component = FakeComponent()
            added.append(component)
            return component

    class FakeProject:
        VBComponents = FakeComponents()

    install_daily_input_journal(FakeProject())
    assert len(added) == 1
    assert added[0].Name == JOURNAL_MODULE_NAME
    assert added[0].CodeModule.source == JOURNAL_MODULE_CODE


def test_vba_source_uses_observed_vbe_canonical_case_for_fullname_references():
    references = [
        line.strip()
        for line in JOURNAL_MODULE_CODE.splitlines()
        if 'QuoteJournalArg(ThisWorkbook.' in line and ' --input ' in line
    ]
    assert len(references) == 2
    assert all(
        line == '" --input " & QuoteJournalArg(ThisWorkbook.fullName) & _'
        for line in references
    )
