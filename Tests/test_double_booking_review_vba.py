from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE = REPO_ROOT / "Python" / "rehab_excel" / "double_booking_review_vba.py"


def test_double_booking_form_hides_technical_ids_and_requires_warning_confirmation():
    source = SOURCE.read_text(encoding="utf-8")
    assert 'DOUBLE_BOOKING_FORM_NAME = "frmDoubleBookingReview"' in source
    assert 'DOUBLE_BOOKING_MODULE_NAME = "modDoubleBookingReview"' in source
    assert 'DOUBLE_BOOKING_MACRO_NAME = "ShowDoubleBookingReview"' in source
    assert 'lstOptions.ColumnWidths = "0 pt;150 pt;60 pt;0 pt;150 pt;60 pt;0 pt"' in source
    assert "Θέλετε να συνεχίσετε;" in source
    assert "vbYesNo + vbExclamation + vbDefaultButton2" in source


def test_double_booking_vba_uses_review_export_and_confirmed_apply_tools():
    source = SOURCE.read_text(encoding="utf-8")
    assert 'ResolveProjectTool("build_double_booking_review_payload.py")' in source
    assert 'ResolveProjectTool("export_double_booking_review_rows.py")' in source
    assert 'ResolveProjectTool("apply_confirmed_double_booking.py")' in source
    assert '" --confirm --overwrite"' in source
    assert "ThisWorkbook.FullName" in source


def test_double_booking_python_processes_are_hidden_and_show_status_progress():
    source = SOURCE.read_text(encoding="utf-8")
    assert "Private Function RunHidden(ByVal commandLine As String) As Long" in source
    assert "shell.Run(commandLine, 0, True)" in source
    assert ".Exec(commandLine)" not in source
    assert 'Application.StatusBar = "Έλεγχος εκκρεμών συνεδριών..."' in source
    assert 'Application.StatusBar = "Προετοιμασία επιλογών διπλοκράτησης..."' in source
    assert 'Application.StatusBar = "Εφαρμογή επιβεβαιωμένης διπλοκράτησης..."' in source
    assert "Application.StatusBar = False" in source


def test_ui_slice_is_preview_only_not_authoritative_commit():
    source = SOURCE.read_text(encoding="utf-8")
    assert "DOUBLE_BOOKING_CONFIRMED_UI_PREVIEW.xlsm" in source
    assert "authoritative_commit_worker_cli.py" not in source
    assert "StartAuthoritativeCommit" not in source
