from rehab_excel.patient_history_report_vba import PATIENT_REPORT_VBA, install_patient_report_components
import inspect


def test_patient_history_unmerges_before_clear_on_mode_changes():
    source = PATIENT_REPORT_VBA
    assert source.index('ws.Range("A1:I2").UnMerge') < source.index('ws.Range("A1:I1048576").ClearContents')
    assert source.index('ws.Unprotect') < source.index('ws.Range("A1:I2").UnMerge')
    assert 'ws.Range("A1:I1").Merge' in source
    assert 'ws.Range("A2:I2").Merge' in source


def test_patient_history_report_error_contains_vba_number_and_description():
    source = PATIENT_REPORT_VBA
    assert "failureNumber = Err.Number" in source
    assert "failureDescription = Err.Description" in source
    assert 'CStr(failureNumber)' in source
    assert 'Application.EnableEvents = priorEvents' in source
    assert 'Application.ScreenUpdating = priorScreenUpdating' in source


def test_companion_statistics_buttons_have_matching_geometry_and_readable_font():
    source = inspect.getsource(install_patient_report_components)
    assert 'button.Width = 146' in source
    assert 'button.Height = 26' in source
    assert 'shape = source.Shapes.AddShape(5, left, top, 146, 26)' in source
    assert 'left = base_left + 154' in source
    assert 'button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0' in source
    assert 'shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0' in source
    assert 'button.TextFrame2.TextRange.Font.Fill.Transparency = 0' in source
    assert 'shape.TextFrame2.TextRange.Font.Fill.Transparency = 0' in source
