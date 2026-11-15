"""Synthetic contract for guarded DEMO patient discharge/finish UI."""
from types import SimpleNamespace

from rehab_excel import trial_demo_patient_exit_vba as action
from rehab_excel.trial_demo_weekly_navigation import MASTER_BUTTON as WEEKLY_BUTTON


class FakeCodeModule:
    def __init__(self):
        self.source = ""

    @property
    def CountOfLines(self):
        return len(self.source.splitlines())

    def AddFromString(self, code):
        self.source = code

    def DeleteLines(self, start, count):
        assert start == 1 and count == self.CountOfLines
        self.source = ""


class FakeComponent:
    def __init__(self, collection):
        self._collection = collection
        self.CodeModule = FakeCodeModule()
        self._name = None

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, name):
        self._collection.entries[name] = self
        self._name = name


class FakeComponents:
    def __init__(self):
        self.entries = {}

    def __call__(self, name):
        return self.entries[name]

    def Add(self, kind):
        assert kind == 1
        return FakeComponent(self)


class FakeMaster:
    def Range(self, key):
        assert key == "A2"
        return SimpleNamespace(Left=10, Top=12)


class FakeWorkbook:
    def __init__(self):
        self.VBProject = SimpleNamespace(VBComponents=FakeComponents())
        self.master = FakeMaster()

    def Worksheets(self, name):
        assert name == "MASTER_SCHEDULE"
        return self.master


def test_native_demo_exit_button_wires_to_isolated_module_without_toolbar_mutation(monkeypatch):
    calls = []
    monkeypatch.setattr(action, "_shape", lambda sheet, **kwargs: calls.append(kwargs))
    wb = FakeWorkbook()
    action.install_demo_patient_exit_action(wb)
    assert len(calls) == 1
    assert calls[0]["name"] == action.BUTTON_NAME
    assert calls[0]["action"] == action.MACRO_NAME
    assert calls[0]["left"] > 500
    assert calls[0]["name"] != WEEKLY_BUTTON
    assert wb.VBProject.VBComponents(action.MODULE_NAME).CodeModule.source == action.VBA_CODE


def test_native_demo_exit_action_install_is_repeatable(monkeypatch):
    calls = []
    monkeypatch.setattr(action, "_shape", lambda sheet, **kwargs: calls.append(kwargs))
    wb = FakeWorkbook()
    action.install_demo_patient_exit_action(wb)
    action.install_demo_patient_exit_action(wb)
    assert len(wb.VBProject.VBComponents.entries) == 1
    assert len(calls) == 2
    assert wb.VBProject.VBComponents(action.MODULE_NAME).CodeModule.source == action.VBA_CODE


def test_exit_is_demo_only_and_explicitly_user_confirmed():
    code = action.VBA_CODE
    assert 'ThisWorkbook.Name, "_DEMO_"' in code
    assert "If ThisWorkbook.ReadOnly Then" in code
    assert 'Set exits = ThisWorkbook.Worksheets(EXIT_SHEET)' in code
    assert "If Not LedgerReady(exits, patients) Then" in code
    assert "PatientRowForId(patients, patientId)" in code
    assert "If Not ParseExactIsoDate(dateInput, serviceEnd) Then" in code
    assert "vbYesNo + vbDefaultButton2" in code
    assert "If answer <> vbYes Then Exit Sub" in code


def test_vba_preserves_histories_and_writes_only_a_date_effective_event():
    code = action.VBA_CODE
    assert 'exits.Cells(exitRow, 1).Value = patientId' in code
    assert 'exits.Cells(exitRow, 2).Value = exitKind' in code
    assert 'exits.Cells(exitRow, 3).Value = serviceEnd' in code
    assert 'exits.Cells(exitRow, 4).Value = reason' in code
    assert "ThisWorkbook.Save" in code
    assert "If written Then" in code
    assert "ClearContents" in code
    assert "Application.EnableEvents = oldEvents" in code
    assert "Application.ScreenUpdating = oldScreen" in code
    for unsafe in ("PATIENTS.Delete", "PATIENT_PLANNER.Delete", "ClearAllHistories",
                   "Worksheets(\"SESSIONS\").Delete", "ThisWorkbook.SaveAs"):
        assert unsafe not in code


def test_vba_refuses_duplicate_identity_and_wrong_patient_exit_kind():
    code = action.VBA_CODE
    assert 'FindTypeColumn(patients)' in code
    assert '"inpatient_discharge"' in code
    assert '"outpatient_program_end"' in code
    assert "matches <> 1" in code
    assert "If kind <> ExpectedExitKind" in code
    assert "For prior = 2 To rowNumber - 1" in code
    assert "StrComp(normalizedId, normalizedPrior, vbTextCompare)" in code


def test_seed_installer_and_native_readback_are_both_connected():
    import inspect
    from rehab_excel import trial_demo_native_seed as seed
    backend = inspect.getsource(seed.Win32ComTrialDemoSeedBackend.write)
    command = inspect.getsource(seed.seed_trial_demo_workbook)
    assert "install_demo_patient_exit_action(wb)" in backend
    assert "verify_demo_patient_exit_action(stage)" in command
