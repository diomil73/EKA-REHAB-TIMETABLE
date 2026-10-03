from datetime import date, time

from rehab_excel.daily_input_sheet import (
    PATIENT_CANCELLATION_STATUSES,
    _format_daily_sheet,
    _get_or_create_sheet,
    build_daily_input_spec,
)


def test_daily_input_spec_keeps_existing_statuses_and_appends_cancellation_options():
    statuses = ("ΠΑΡΩΝ", "εκτός κλινικής", "Αναβολή ΦΘ")
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α", "Β"),
        patient_names=("ΑΣΘΕΝΗΣ 1", "ΑΣΘΕΝΗΣ 2"),
        patient_statuses=statuses,
        timeslots=(time(8, 30), time(9, 15)),
    )
    assert spec.patient_statuses == (*statuses, *PATIENT_CANCELLATION_STATUSES)


def test_daily_input_spec_deduplicates_dropdown_values_without_reordering():
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α", "Α", "Β"),
        patient_names=("Π1", "Π2", "Π1"),
        patient_statuses=(
            "ΠΑΡΩΝ",
            "ΠΑΡΩΝ",
            "ΑΠΩΝ",
            "ΑΝΑΒΟΛΗ ΤΜΗΜΑΤΟΣ",
        ),
        timeslots=(time(8, 30), time(9, 15)),
    )
    assert spec.therapist_names == ("Α", "Β")
    assert spec.patient_names == ("Π1", "Π2")
    assert spec.patient_statuses == (
        "ΠΑΡΩΝ",
        "ΑΠΩΝ",
        "ΑΝΑΒΟΛΗ ΤΜΗΜΑΤΟΣ",
        "ΔΕΝ ΠΡΟΣΗΛΘΕ",
    )


def test_daily_input_layout_has_separate_therapist_and_patient_sections():
    spec = build_daily_input_spec(
        target_date=date(2026, 9, 19),
        therapist_names=("Α",),
        patient_names=("Π1",),
        patient_statuses=("ΠΑΡΩΝ",),
        timeslots=(time(8, 30),),
        therapist_rows=10,
        patient_rows=20,
    )
    assert spec.therapist_first_row == 7
    assert spec.therapist_last_row == 16
    assert spec.patient_header_row == 20
    assert spec.patient_first_row == 21
    assert spec.patient_last_row == 40



class _FakeCells:
    def __init__(self):
        self.unmerged = False
        self.cleared = False

    def UnMerge(self):
        self.unmerged = True

    def Clear(self):
        self.cleared = True


class _FakeSheet:
    def __init__(self, name):
        self.Name = name
        self.Cells = _FakeCells()


class _FakeWorksheets:
    def __init__(self, sheets):
        self._sheets = sheets

    def __iter__(self):
        return iter(self._sheets)

    def __call__(self, index):
        return self._sheets[index - 1]

    @property
    def Count(self):
        return len(self._sheets)

    def Add(self, **kwargs):
        sheet = _FakeSheet("SheetX")
        if "Before" in kwargs:
            self._sheets.insert(0, sheet)
        else:
            self._sheets.append(sheet)
        return sheet


class _FakeWorkbook:
    def __init__(self, sheets):
        self.Worksheets = _FakeWorksheets(sheets)


def test_get_or_create_sheet_reuses_existing_sheet_without_delete():
    existing = _FakeSheet("DAILY_INPUT")
    workbook = _FakeWorkbook([existing])

    result = _get_or_create_sheet(workbook, "DAILY_INPUT", before_first=True)

    assert result is existing
    assert existing.Cells.unmerged is True
    assert existing.Cells.cleared is True
    assert workbook.Worksheets.Count == 1


def test_get_or_create_sheet_creates_missing_sheet_in_requested_position():
    workbook = _FakeWorkbook([_FakeSheet("SETTINGS")])

    result = _get_or_create_sheet(workbook, "DAILY_INPUT", before_first=True)

    assert result.Name == "DAILY_INPUT"
    assert workbook.Worksheets(1) is result



def test_daily_input_sheet_contains_outpatient_marker_rule():
    import inspect

    source = inspect.getsource(_format_daily_sheet)
    assert 'LEFT(A{pr1},2)="ΕΞ"' in source
    assert "OUTPATIENT_LIGHT_BLUE_RGB" in source
