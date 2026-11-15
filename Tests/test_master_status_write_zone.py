import pytest

from rehab_excel.layout_contract import LayoutSafetyError, assert_cell_write_allowed


def test_master_daily_status_overlay_may_write_only_current_status_column():
    assert_cell_write_allowed("MASTER_SCHEDULE", "M7")

    with pytest.raises(LayoutSafetyError):
        assert_cell_write_allowed("MASTER_SCHEDULE", "K7")

    with pytest.raises(LayoutSafetyError):
        assert_cell_write_allowed("MASTER_SCHEDULE", "L7")


def test_master_therapy_display_zone_remains_allowed():
    assert_cell_write_allowed("MASTER_SCHEDULE", "D7")
    assert_cell_write_allowed("MASTER_SCHEDULE", "J5000")
