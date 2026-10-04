import inspect

from rehab_excel.application_shell_vba import (
    APP_SHELL_MODULE_CODE,
    APP_SHELL_MODULE_NAME,
    _ensure_event_call,
    _workbook_document_component,
)
from rehab_excel.registration_menu_vba import Win32ComRegistrationMenuInstaller


class FakeCodeModule:
    def __init__(self, text=""):
        self.lines = text.splitlines()

    @property
    def CountOfLines(self):
        return len(self.lines)

    def Lines(self, start, count):
        return "\r\n".join(self.lines[start - 1 : start - 1 + count])

    def ProcStartLine(self, procedure_name, _kind):
        needle = f"sub {procedure_name.lower()}("
        for index, line in enumerate(self.lines, start=1):
            if needle in line.strip().lower():
                return index
        raise RuntimeError("procedure not found")

    def ProcCountLines(self, procedure_name, kind):
        start = self.ProcStartLine(procedure_name, kind)
        for index in range(start, len(self.lines) + 1):
            if self.lines[index - 1].strip().lower() == "end sub":
                return index - start + 1
        raise RuntimeError("End Sub not found")

    def InsertLines(self, line_number, text):
        additions = text.replace("\r\n", "\n").split("\n")
        self.lines[line_number - 1 : line_number - 1] = additions


def test_application_shell_targets_master_and_hides_excel_chrome():
    assert APP_SHELL_MODULE_NAME == "modApplicationShell"
    assert 'Worksheets("MASTER_SCHEDULE").Activate' in APP_SHELL_MODULE_CODE
    assert "Application.DisplayFormulaBar = False" in APP_SHELL_MODULE_CODE
    assert "Application.DisplayStatusBar = False" in APP_SHELL_MODULE_CODE
    assert 'SetRibbonVisible False' in APP_SHELL_MODULE_CODE
    assert "ActiveWindow.DisplayWorkbookTabs = False" in APP_SHELL_MODULE_CODE
    assert "ActiveWindow.DisplayHeadings = False" in APP_SHELL_MODULE_CODE
    assert "ActiveWindow.DisplayGridlines = False" in APP_SHELL_MODULE_CODE


def test_application_shell_has_explicit_excel_restore_escape_hatch():
    assert "Public Sub ExitApplicationShell()" in APP_SHELL_MODULE_CODE
    assert "Public Sub RestoreExcelInterface()" in APP_SHELL_MODULE_CODE
    assert "Application.DisplayFormulaBar = previousFormulaBar" in APP_SHELL_MODULE_CODE
    assert "Application.DisplayStatusBar = previousStatusBar" in APP_SHELL_MODULE_CODE
    assert "SetRibbonVisible previousRibbonVisible" in APP_SHELL_MODULE_CODE


def test_existing_workbook_open_is_extended_not_replaced():
    module = FakeCodeModule(
        "Option Explicit\n"
        "Private Sub Workbook_Open()\n"
        "    ExistingStartupCall\n"
        "End Sub"
    )

    _ensure_event_call(
        module,
        procedure_name="Workbook_Open",
        signature="Private Sub Workbook_Open()",
        call_name="EnterApplicationShell",
    )

    text = "\n".join(module.lines)
    assert "ExistingStartupCall" in text
    assert text.count("EnterApplicationShell") == 1


def test_missing_before_close_event_is_created():
    module = FakeCodeModule("Option Explicit")

    _ensure_event_call(
        module,
        procedure_name="Workbook_BeforeClose",
        signature="Private Sub Workbook_BeforeClose(Cancel As Boolean)",
        call_name="ExitApplicationShell",
    )

    text = "\n".join(module.lines)
    assert "Private Sub Workbook_BeforeClose(Cancel As Boolean)" in text
    assert "ExitApplicationShell" in text
    assert "End Sub" in text


def test_event_call_installation_is_idempotent():
    module = FakeCodeModule(
        "Private Sub Workbook_Activate()\n"
        "    EnterApplicationShell\n"
        "End Sub"
    )

    _ensure_event_call(
        module,
        procedure_name="Workbook_Activate",
        signature="Private Sub Workbook_Activate()",
        call_name="EnterApplicationShell",
    )

    assert "\n".join(module.lines).count("EnterApplicationShell") == 1


def test_registration_menu_installer_installs_application_shell():
    source = inspect.getsource(Win32ComRegistrationMenuInstaller.install)
    assert "install_application_shell(vbproject, workbook=workbook)" in source
    assert "Application shell module was not confirmed" in source



class FakeComponent:
    def __init__(self, name, component_type=100):
        self.Name = name
        self.Type = component_type


class FakeVBComponents:
    def __init__(self, components):
        self._components = list(components)
        self.Count = len(self._components)

    def __call__(self, key):
        if isinstance(key, int):
            return self._components[key - 1]
        for component in self._components:
            if component.Name == key:
                return component
        raise RuntimeError("missing component")


class FakeVBProject:
    def __init__(self, components):
        self.VBComponents = FakeVBComponents(components)


class FakeWorkbook:
    def __init__(self, code_name):
        self.CodeName = code_name


def test_workbook_document_component_uses_actual_workbook_codename():
    localized = FakeComponent("ΒιβλίοΕργασίας")
    project = FakeVBProject([FakeComponent("Sheet1"), localized])
    workbook = FakeWorkbook("ΒιβλίοΕργασίας")

    assert _workbook_document_component(project, workbook) is localized


def test_workbook_document_component_falls_back_to_thisworkbook():
    expected = FakeComponent("ThisWorkbook")
    project = FakeVBProject([FakeComponent("Sheet1"), expected])

    assert _workbook_document_component(project) is expected
