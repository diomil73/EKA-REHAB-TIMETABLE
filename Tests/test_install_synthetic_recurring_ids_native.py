"""Synthetic COM-only checks for immutable assignment migration installer.

No actual Excel, personal data or clinical workbook is used in these tests.
"""
import pytest

from rehab_excel.recurring_id_migration_plan import (
    AssignmentIdRow, RecurringIdMigrationPlan,
)
from tools.install_synthetic_recurring_ids_native import (
    NativeIdInstallError, _install_identity_cells,
)


class Cell:
    def __init__(self, value=None):
        self.Value = value

    @property
    def Row(self):
        return 1

    def End(self, direction):
        assert direction == -4162
        return self


class Cells:
    def __init__(self, sheet):
        self.sheet = sheet

    def __call__(self, row, col):
        return self.sheet.cells.setdefault((row, col), Cell())


class Rows:
    Count = 1048576


class Sheet:
    def __init__(self, name, rows=()):
        self.Name = name
        self.cells = {(row, col): Cell(value) for row, col, value in rows}
        self.Cells = Cells(self)
        self.Rows = Rows()


class Sheets:
    def __init__(self):
        self.items = [Sheet("PROGRAM_WEEKLY_SLOTS", (
            (1, 1, "PatientID"),
            (1, 2, "Treatment"),
            (2, 1, "SYN-P"),
            (2, 2, "ΦΘ"),
        )), Sheet("OUTPATIENT_SCHEDULE", (
            (1, 1, "PatientID"),
            (1, 3, "Θεραπεία"),
            (2, 1, "SYN-O"),
            (2, 3, "Εργο"),
        ))]
        self.added = False

    def __iter__(self):
        return iter(self.items)

    def __call__(self, name):
        return next(s for s in self.items if s.Name == name)

    def Add(self):
        self.added = True
        sheet = Sheet("_NEW")
        self.items.append(sheet)
        return sheet


class Book:
    def __init__(self):
        self.Worksheets = Sheets()


def plan():
    return RecurringIdMigrationPlan((
        AssignmentIdRow("assignment:PLANNER00000001", "planner:5:ΦΘ",
                        "SYN-I", "ΦΘ", "PATIENT_PLANNER", None),
        AssignmentIdRow("assignment:WEEKLY00000001", "weekly:2",
                        "SYN-P", "ΦΘ", "PROGRAM_WEEKLY_SLOTS", 2),
        AssignmentIdRow("assignment:OUTPAT00000001", "outpatient:2",
                        "SYN-O", "Εργο", "OUTPATIENT_SCHEDULE", 2),
    ))


def test_installer_adds_only_identity_surface_without_rewriting_clinical_rows():
    book = Book()
    _install_identity_cells(book, plan())
    reg = book.Worksheets("RECURRING_ASSIGNMENT_IDS")
    assert [reg.Cells(1, i).Value for i in (1, 2, 3)] == [
        "AssignmentID", "PatientID", "Treatment",
    ]
    assert [reg.Cells(2, i).Value for i in (1, 2, 3)] == [
        "assignment:PLANNER00000001", "SYN-I", "ΦΘ",
    ]
    weekly = book.Worksheets("PROGRAM_WEEKLY_SLOTS")
    assert weekly.Cells(2, 1).Value == "SYN-P"
    assert weekly.Cells(2, 2).Value == "ΦΘ"
    assert weekly.Cells(1, 7).Value == "AssignmentID"
    assert weekly.Cells(2, 7).Value == "assignment:WEEKLY00000001"
    outpatient = book.Worksheets("OUTPATIENT_SCHEDULE")
    assert outpatient.Cells(2, 3).Value == "Εργο"
    assert outpatient.Cells(2, 7).Value == "assignment:OUTPAT00000001"


def test_preflight_rejects_changed_row_before_adding_sheet_or_writing_ids():
    book = Book()
    book.Worksheets("PROGRAM_WEEKLY_SLOTS").Cells(2, 1).Value = "OTHER-P"
    with pytest.raises(NativeIdInstallError, match="no longer matches"):
        _install_identity_cells(book, plan())
    assert not book.Worksheets.added
    assert book.Worksheets("PROGRAM_WEEKLY_SLOTS").Cells(1, 7).Value is None


def test_preflight_rejects_existing_immutable_identity_surface():
    book = Book()
    book.Worksheets("OUTPATIENT_SCHEDULE").Cells(1, 7).Value = "AssignmentID"
    with pytest.raises(NativeIdInstallError, match="already occupied"):
        _install_identity_cells(book, plan())
    assert not book.Worksheets.added


def test_preflight_rejects_existing_registry_even_if_empty():
    book = Book()
    book.Worksheets.items.append(Sheet("RECURRING_ASSIGNMENT_IDS"))
    with pytest.raises(NativeIdInstallError, match="already exists"):
        _install_identity_cells(book, plan())
    assert not book.Worksheets.added
