import pytest

from rehab_excel.layout_contract import LayoutSafetyError, assert_cell_write_allowed, write_zones_for_sheet


def test_master_status_column_is_a_narrow_confirmed_write_zone():
    zones = write_zones_for_sheet("MASTER_SCHEDULE")
    pairs = {(zone.start, zone.end) for zone in zones}

    assert ("D2", "J5000") in pairs
    assert ("M2", "M5000") in pairs

    assert_cell_write_allowed("MASTER_SCHEDULE", "M2")
    assert_cell_write_allowed("MASTER_SCHEDULE", "M5000")


def test_master_status_header_and_non_derived_columns_remain_protected():
    for cell in ("M1", "A2", "B7", "C7", "K7", "L7", "N7"):
        with pytest.raises(LayoutSafetyError):
            assert_cell_write_allowed("MASTER_SCHEDULE", cell)
