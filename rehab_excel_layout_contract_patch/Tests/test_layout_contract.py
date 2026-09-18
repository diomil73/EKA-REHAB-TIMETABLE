from datetime import time

import pytest

from rehab_excel.layout_contract import (
    LayoutSafetyError,
    assert_cell_write_allowed,
    cell_in_range,
    resolve_master_schedule_rows_from_cells,
    resolve_therapist_daily_cell_from_cells,
)


def _daily_cells():
    cells = {
        "B1": "Αργέντος",
        "C1": "Βασιλειάδης",
        "D1": "Γαβράς",
        "E1": "Γουνελάς",
        "F1": "Καλυβιώτης",
        "G1": "Κοσμίδης",
        "H1": "Μηλιδάκης",
        "I1": "Παπακίτσου",
        "J1": "Ορφανοι",
        "B11": "Παπαμάνου",
        "C11": "Πέτσιος",
        "D11": "Σάκκου",
        "E11": "Σάρρας",
        "F11": "Σκαρώνης",
        "G11": "Φιλιππούσης",
        "H11": "φοιτ 1",
        "I11": "Χρήστου",
        "J11": "Ψιάχου",
    }
    excel_times = {
        2: 8.5 / 24,
        3: 9.25 / 24,
        4: 10 / 24,
        5: 10.75 / 24,
        6: 11.5 / 24,
        7: 12.25 / 24,
        8: 13 / 24,
    }
    for row, value in excel_times.items():
        cells[f"A{row}"] = str(value)
        cells[f"A{row + 10}"] = str(value)
    return cells


def test_cell_range_helper_handles_normal_grid_cells():
    assert cell_in_range("B2", "B2", "J8") is True
    assert cell_in_range("J8", "B2", "J8") is True
    assert cell_in_range("A2", "B2", "J8") is False


def test_daily_write_zone_allows_patient_grid():
    assert_cell_write_allowed("THERAPIST_DAILY", "H3")
    assert_cell_write_allowed("THERAPIST_DAILY", "D14")


def test_daily_write_zone_protects_headers_times_and_button_area():
    for cell in ("B1", "A3", "K2"):
        with pytest.raises(LayoutSafetyError):
            assert_cell_write_allowed("THERAPIST_DAILY", cell)


def test_master_schedule_write_zone_protects_identity_columns():
    assert_cell_write_allowed("MASTER_SCHEDULE", "D2")
    assert_cell_write_allowed("MASTER_SCHEDULE", "J5000")
    for cell in ("A2", "B2", "C2", "K2"):
        with pytest.raises(LayoutSafetyError):
            assert_cell_write_allowed("MASTER_SCHEDULE", cell)


def test_unmapped_future_sheet_is_still_blocked_at_cell_level():
    with pytest.raises(LayoutSafetyError, match="no confirmed cell-level write zone"):
        assert_cell_write_allowed("STATISTICS", "A1")


def test_resolves_top_block_therapist_daily_cell():
    cell = resolve_therapist_daily_cell_from_cells(
        _daily_cells(), "Μηλιδάκης", time(9, 15)
    )
    assert cell == "H3"


def test_resolves_bottom_block_therapist_daily_cell():
    cell = resolve_therapist_daily_cell_from_cells(
        _daily_cells(), "Σάκκου", time(10, 0)
    )
    assert cell == "D14"


def test_resolves_student_daily_cell_as_normal_provider_column():
    cell = resolve_therapist_daily_cell_from_cells(
        _daily_cells(), "φοιτ 1", time(12, 15)
    )
    assert cell == "H17"


def test_daily_cell_resolution_is_name_normalized():
    cell = resolve_therapist_daily_cell_from_cells(
        _daily_cells(), "  ΜΗΛΙΔΑΚΗΣ  ", time(9, 15)
    )
    assert cell == "H3"


def test_daily_cell_resolution_rejects_unknown_provider():
    with pytest.raises(LayoutSafetyError, match="No THERAPIST_DAILY cell"):
        resolve_therapist_daily_cell_from_cells(
            _daily_cells(), "ΑΓΝΩΣΤΟΣ", time(9, 15)
        )


def test_master_schedule_patient_lookup_returns_all_matching_rows():
    cells = {
        "C2": "ΖΑΛΟΚΩΣΤΑΣ",
        "C3": "ΠΑΠΑΔΟΠΟΥΛΟΥ",
        "C10": " ΖΑΛΟΚΩΣΤΑΣ ",
    }
    assert resolve_master_schedule_rows_from_cells(cells, "ζαλοκωστας") == (2, 10)
