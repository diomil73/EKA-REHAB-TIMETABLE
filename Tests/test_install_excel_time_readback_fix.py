import inspect

from tools import install_excel_time_readback_fix as module


def test_fix_targets_both_authoritative_vba_modules():
    assert module.MODULES == ("modCloseDayAction", "modDayAmendmentAction")


def test_fix_accepts_numeric_excel_time_serials():
    assert 'If IsIsoTimeText(expectedText) Then' in module.NEW_BLOCK
    assert 'If IsNumeric(actualValue) Then' in module.NEW_BLOCK
    assert 'Format$(CDbl(actualValue), "hh:nn:ss") = expectedText' in module.NEW_BLOCK


def test_fix_reopens_workbook_to_verify_persisted_vba():
    source = inspect.getsource(module.install_excel_time_readback_fix)
    assert 'workbook.Close(SaveChanges=False)' in source
    assert 'excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)' in source
    assert '_verify(workbook)' in source


def test_fix_rejects_legacy_date_only_comparator_after_patch():
    source = inspect.getsource(module._verify)
    assert 'if OLD_LINE in source' in source
    assert 'Legacy Date-only time comparator still present' in source
