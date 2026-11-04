import inspect

from tools import audit_workbook_protection as module


def test_excel_column_letters_cover_single_and_multiple_letter_columns():
    assert [module.column_letter(n) for n in (1, 26, 27, 52, 53, 702, 703)] == [
        "A", "Z", "AA", "AZ", "BA", "ZZ", "AAA"
    ]


def test_unlocked_cell_ranges_are_compressed_per_row():
    assert module.compress_unlocked_columns(12, [1, 2, 3, 5, 8, 9]) == [
        "A12:C12", "E12", "H12:I12"
    ]
    assert module.compress_unlocked_columns(4, []) == []


def test_excel_selection_permissions_are_described():
    assert module.selection_label(0) == "xlNoRestrictions"
    assert module.selection_label(1) == "xlUnlockedCells"
    assert module.selection_label(2) == "xlNoSelection"
    assert module.selection_label(99) == "unknown(99)"


def test_protection_audit_is_read_only_and_never_changes_sheet_state():
    source = inspect.getsource(module.audit)
    assert "ReadOnly=True" in source
    assert "UpdateLinks=0" in source
    assert "workbook.Close(SaveChanges=False)" in source
    assert ".Protect(" not in source
    assert ".Unprotect(" not in source
    assert ".Save(" not in source


def test_protection_audit_reports_contents_drawing_and_structure_protection():
    source = inspect.getsource(module.audit)
    assert '"ProtectStructure"' in source
    assert '"ProtectWindows"' in source
    assert '"ProtectContents"' in source
    assert '"ProtectDrawingObjects"' in source
    assert '"EnableSelection"' in source
    assert '"UsedRange.Locked"' in source
