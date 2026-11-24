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


def test_double_booking_vba_uses_combined_ui_prep_and_confirmed_apply_tools():
    source = SOURCE.read_text(encoding="utf-8")
    assert 'ResolveProjectTool("prepare_double_booking_review_ui.py")' in source
    assert 'ResolveProjectTool("apply_confirmed_double_booking.py")' in source
    assert 'ResolveProjectTool("build_double_booking_review_payload.py")' not in source
    assert 'ResolveProjectTool("export_double_booking_review_rows.py")' not in source
    assert '" --confirm --overwrite"' in source
    assert '" --response " & QuoteArg(responsePath)' in source
    assert "ThisWorkbook.FullName" in source


def test_double_booking_python_processes_are_hidden_and_show_status_progress():
    source = SOURCE.read_text(encoding="utf-8")
    assert "Private Function RunHidden(ByVal commandLine As String) As Long" in source
    assert "shell.Run(commandLine, 0, True)" in source
    assert ".Exec(commandLine)" not in source
    assert 'Application.StatusBar = "Έλεγχος εκκρεμών συνεδριών και προετοιμασία επιλογών..."' in source
    assert 'Application.StatusBar = "Εφαρμογή και επαλήθευση διπλοκράτησης..."' in source
    assert "Application.StatusBar = False" in source


def test_double_booking_ui_commits_verified_preview_and_reopens_authoritative_workbook():
    source = SOURCE.read_text(encoding="utf-8")
    assert "DOUBLE_BOOKING_CONFIRMED_UI_PREVIEW.xlsm" in source
    assert 'ResolveProjectTool("authoritative_commit_worker_cli.py")' in source
    assert "StartAuthoritativeCommit(previewPath, sourceHash)" in source
    assert 'JsonStringValue(responseText, "source_sha256_before")' in source
    assert 'JsonStringValue(responseText, "output_path")' in source
    assert 'q & "remove_preview_after_success" & q & ":true,"' in source
    assert 'q & "reopen" & q & ":true,"' in source
    assert 'q & "post_reopen_target" & q & ":" & q & "REPLACEMENTS" & q' in source
    assert "CreateObject(\"WScript.Shell\").Run commandLine, 0, False" in source
    assert "ThisWorkbook.Close SaveChanges:=False" in source

def test_double_booking_confirmed_apply_cannot_bypass_disabled_journal():
    from rehab_excel.double_booking_review_vba import MODULE_CODE
    apply = MODULE_CODE.split("Public Function ApplySelectedDoubleBooking", 1)[1].split(
        "End Function", 1
    )[0]
    guard = "If Not DailyInputJournalApplyEnabled() Then"
    assert guard in apply
    assert apply.index(guard) < apply.index("apply_confirmed_double_booking.py")
    assert apply.index(guard) < apply.index("StartAuthoritativeCommit(previewPath, sourceHash)")
    assert apply.index(guard) < apply.index("ThisWorkbook.Close")
    assert "Exit Function" in apply.split(guard, 1)[1].split("End If", 1)[0]
    assert "Δεν έγιναν αλλαγές" in apply


def test_internal_legacy_source_commit_worker_is_also_guarded():
    from rehab_excel.double_booking_review_vba import MODULE_CODE
    worker = MODULE_CODE.split("Private Function StartAuthoritativeCommit", 1)[1].split(
        "End Function", 1
    )[0]
    guard = "If Not DailyInputJournalApplyEnabled() Then Exit Function"
    assert guard in worker
    assert worker.index(guard) < worker.index("authoritative_commit_worker_cli.py")
    assert worker.index(guard) < worker.index('CreateObject("WScript.Shell").Run commandLine, 0, False')

    from rehab_excel.daily_input_journal_vba import JOURNAL_MODULE_CODE
    gate = JOURNAL_MODULE_CODE.split(
        "Public Function DailyInputJournalApplyEnabled()", 1
    )[1].split("End Function", 1)[0]
    assert "DailyInputJournalApplyEnabled = False" in gate
