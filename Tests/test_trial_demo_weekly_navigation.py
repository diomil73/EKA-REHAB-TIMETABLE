"""Synthetic only: the private DEMO adds its view to the known Excel toolbar."""
from rehab_excel.trial_demo_weekly_navigation import (
    BACK_BUTTON, BACK_MACRO, DEMO_NAV_MODULE, MASTER_BUTTON,
    OPEN_MACRO, VBA_CODE, install_demo_weekly_navigation,
)


class _Node:
    def __getattr__(self, name):
        child = _Node()
        setattr(self, name, child)
        return child


class _Shape:
    def __init__(self, collection):
        self._collection = collection
        self._name = ""
        self.OnAction = ""
        self.TextFrame2 = _Node()
        self.Fill = _Node()
        self.Line = _Node()

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, name):
        self._name = name
        self._collection.items[name] = self

    def Delete(self):
        self._collection.items.pop(self.Name)


class _Shapes:
    def __init__(self):
        self.items = {}

    def __call__(self, name):
        return self.items[name]

    def AddShape(self, kind, left, top, width, height):
        assert kind == 5 and width > 0 and height > 0
        return _Shape(self)


class _Sheet:
    def __init__(self):
        self.Shapes = _Shapes()

    def Range(self, name):
        assert name in ("A2", "I1")
        anchor = _Node()
        anchor.Left = 10
        anchor.Top = 12
        return anchor


class _CodeModule:
    def __init__(self):
        self.lines = ""

    @property
    def CountOfLines(self):
        return self.lines.count("\n") + bool(self.lines)

    def DeleteLines(self, first, count):
        assert first == 1
        self.lines = ""

    def AddFromString(self, value):
        self.lines += value


class _VbComponent:
    def __init__(self, collection):
        self._collection = collection
        self._name = ""
        self.CodeModule = _CodeModule()

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, name):
        self._name = name
        self._collection.items[name] = self


class _Components:
    def __init__(self):
        self.items = {}

    def __call__(self, name):
        return self.items[name]

    def Add(self, kind):
        assert kind == 1
        return _VbComponent(self)


class _Book:
    def __init__(self):
        self.sheets = {
            "MASTER_SCHEDULE": _Sheet(),
            "DEMO_WEEKLY_VIEW": _Sheet(),
        }
        self.VBProject = _Node()
        self.VBProject.VBComponents = _Components()

    def Worksheets(self, name):
        return self.sheets[name]


def test_demo_buttons_are_added_without_changing_original_toolbar():
    book = _Book()
    original = _Shape(book.sheets["MASTER_SCHEDULE"].Shapes)
    original.Name = "ekaToolbar_Absences"
    original.OnAction = "ToolbarAbsences"

    install_demo_weekly_navigation(book)
    master = book.sheets["MASTER_SCHEDULE"].Shapes.items
    weekly = book.sheets["DEMO_WEEKLY_VIEW"].Shapes.items
    assert set(master) == {"ekaToolbar_Absences", MASTER_BUTTON}
    assert set(weekly) == {BACK_BUTTON}
    assert master[MASTER_BUTTON].OnAction == OPEN_MACRO
    assert weekly[BACK_BUTTON].OnAction == BACK_MACRO
    assert master["ekaToolbar_Absences"] is original
    assert master["ekaToolbar_Absences"].OnAction == "ToolbarAbsences"
    assert book.VBProject.VBComponents(DEMO_NAV_MODULE).CodeModule.lines == VBA_CODE


def test_navigation_install_is_idempotent_on_one_demo_candidate():
    book = _Book()
    install_demo_weekly_navigation(book)
    install_demo_weekly_navigation(book)
    assert list(book.VBProject.VBComponents.items) == [DEMO_NAV_MODULE]
    assert set(book.sheets["MASTER_SCHEDULE"].Shapes.items) == {MASTER_BUTTON}
    assert set(book.sheets["DEMO_WEEKLY_VIEW"].Shapes.items) == {BACK_BUTTON}
    assert book.VBProject.VBComponents(DEMO_NAV_MODULE).CodeModule.lines == VBA_CODE


def test_weekly_navigation_respects_hidden_tabs_and_restores_application_shell():
    assert 'Worksheets("DEMO_WEEKLY_VIEW").Visible = xlSheetVisible' in VBA_CODE
    assert 'Worksheets("DEMO_WEEKLY_VIEW").Activate' in VBA_CODE
    assert "KeepApplicationShell" in VBA_CODE
    assert "Public Sub OpenDemoWeeklyView()" in VBA_CODE
