import inspect

from rehab_excel.outpatient_schedule_sheet import (
    OUTPATIENT_SCHEDULE_HEADERS,
    OUTPATIENT_SCHEDULE_SHEET,
    ensure_outpatient_schedule_sheet_in_workbook,
)
from tools import build_outpatient_schedule_form_preview


def test_outpatient_schedule_sheet_contract_is_stable():
    assert OUTPATIENT_SCHEDULE_SHEET == "OUTPATIENT_SCHEDULE"
    assert OUTPATIENT_SCHEDULE_HEADERS == (
        "PatientID",
        "Ασθενής",
        "Θεραπεία",
        "Ώρα",
        "Ημέρες",
        "Θεραπευτής",
    )


def test_initializer_is_build_time_excel_com_and_schema_fail_closed():
    source = inspect.getsource(ensure_outpatient_schedule_sheet_in_workbook)

    assert 'DispatchEx("Excel.Application")' in source
    assert "ReadOnly=False" in source
    assert "workbook.Worksheets.Add" in source
    assert "OUTPATIENT_SCHEDULE_HEADERS" in source
    assert "unexpected header schema" in source
    assert "workbook.Save()" in source


def test_canonical_outpatient_build_initializes_source_before_manager_install():
    source = inspect.getsource(build_outpatient_schedule_form_preview.main)

    ensure_pos = source.index("ensure_outpatient_schedule_sheet_in_workbook(report.output_path)")
    edit_pos = source.index("install_outpatient_schedule_edit_support_in_workbook(report.output_path)")
    manager_pos = source.index("install_outpatient_schedule_manager_in_workbook(report.output_path)")

    assert ensure_pos < edit_pos < manager_pos
    assert 'print("Outpatient schedule source: verified")' in source
