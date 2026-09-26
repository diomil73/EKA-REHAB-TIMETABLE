from rehab_excel.native_excel import Win32ComExcelBackend
from rehab_excel.writeback import CellPatch, WriteIntent


class FakeInterior:
    def __init__(self):
        self.Color = None
        self.Pattern = None


class FakeFont:
    Size = 11
    Strikethrough = False
    Italic = False
    Color = None


class FakeCell:
    def __init__(self):
        self.Value = "KEEP-ME"
        self.WrapText = False
        self.Font = FakeFont()
        self.Interior = FakeInterior()

    def Borders(self, _edge):
        raise AssertionError("Borders should not be touched in this test")


class FakeSheet:
    def __init__(self, cell):
        self.cell = cell

    def Range(self, _address):
        return self.cell


class FakeWorksheets:
    def __init__(self, sheet):
        self.sheet = sheet

    def __call__(self, _name):
        return self.sheet


class FakeWorkbook:
    def __init__(self, cell):
        self.Worksheets = FakeWorksheets(FakeSheet(cell))


def test_presentation_patch_does_not_overwrite_cell_value():
    cell = FakeCell()
    backend = Win32ComExcelBackend()
    patch = CellPatch(
        "THERAPIST_DAILY",
        "B2",
        intent=WriteIntent.PRESENTATION,
        fill_role="outpatient_light_blue",
    )

    backend._apply_patch(FakeWorkbook(cell), patch)

    assert cell.Value == "KEEP-ME"
    assert cell.Interior.Color == backend.palette.outpatient_light_blue


def test_clear_fill_is_presentation_only_and_keeps_value():
    cell = FakeCell()
    backend = Win32ComExcelBackend()
    patch = CellPatch(
        "THERAPIST_DAILY",
        "B2",
        intent=WriteIntent.PRESENTATION,
        fill_role="clear_fill",
    )

    backend._apply_patch(FakeWorkbook(cell), patch)

    assert cell.Value == "KEEP-ME"
    assert cell.Interior.Pattern == -4142
