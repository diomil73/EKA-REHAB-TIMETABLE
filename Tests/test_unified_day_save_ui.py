from pathlib import Path

from tools.install_unified_day_save_ui import (
    DAILY_BUTTON,
    MASTER_BUTTON,
    OLD_BUTTONS,
    SOURCE,
)


def test_single_entrypoint_routes_to_existing_authoritative_actions():
    assert "Public Sub SaveOperationalDay()" in SOURCE
    assert "CloseOperationalDay" in SOURCE
    assert "StartClosedDayAmendment" in SOURCE
    assert 'Range("B2")' in SOURCE
    assert 'Worksheets("FINALIZED_DAY_SNAPSHOTS")' in SOURCE
    assert 'status = "CURRENT_FINAL"' in SOURCE
    assert "Case Else" in SOURCE
    assert "GoTo SafetyStop" in SOURCE


def test_installer_has_persisted_reopen_verification_and_deduplicates_buttons():
    script = Path(__file__).resolve().parents[1] / "Python" / "tools" / "install_unified_day_save_ui.py"
    text = script.read_text(encoding="utf-8")
    assert set(OLD_BUTTONS) == {"ekaToolbar_CloseDay", "ekaToolbar_DayAmendment"}
    assert DAILY_BUTTON != MASTER_BUTTON
    assert "_delete_if_present(daily, name)" in text
    assert 'workbook.Worksheets("MASTER_SCHEDULE")' in text
    assert 'Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)' in text
    assert "Persisted VBA reopen verification: OK" in text


def test_comparator_fix_must_be_installed_separately_before_native_save():
    # Routing does not attempt to change the authoritative persistence modules.
    assert "Patient history read-back" not in SOURCE
    assert "PersistDaySavePayload" not in SOURCE
    assert "PersistCloseDayPayload" not in SOURCE
