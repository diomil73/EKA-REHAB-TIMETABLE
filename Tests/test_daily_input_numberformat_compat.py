from datetime import date
from pathlib import Path


def test_daily_input_date_is_written_as_text_without_numberformat_assignment():
    source = Path('Python/rehab_excel/daily_input_sheet.py').read_text(encoding='utf-8')
    assert 'spec.target_date.strftime("%d/%m/%Y")' in source
    assert 'Range("B2").NumberFormat' not in source
