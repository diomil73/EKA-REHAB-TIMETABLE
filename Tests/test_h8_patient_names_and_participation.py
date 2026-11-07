from datetime import date
from types import SimpleNamespace

import pytest

from tools import export_patient_history_ui as ui
from rehab_excel.patient_history_report_vba import PATIENT_REPORT_VBA


def _patient(patient_id, display_name):
    return SimpleNamespace(patient_id=patient_id, display_name=display_name)


def _row(patient_id, *, scheduled=4, performed=3):
    return SimpleNamespace(
        period_start=date(2026, 10, 30),
        period_end=date(2026, 10, 30),
        patient_id=patient_id,
        finalized_days=1,
        scheduled=scheduled,
        performed=performed,
        patient_absences=scheduled - performed,
        cancellations=0,
        replacements=0,
    )


def test_identity_matches_zero_padded_registry_keys():
    assert ui._identity("01") == ui._identity("1")
    assert ui._identity("P01") != ui._identity("01")


def test_patient_registry_duplicate_numeric_identity_is_a_safety_stop():
    with pytest.raises(ValueError, match="Duplicate patient identifier"):
        ui._patient_names([_patient("1", "Α"), _patient("01", "Β")])


def test_participation_denominator_and_empty_state():
    assert ui._completion_pct(3, 4) == "75.0"
    assert ui._completion_pct(0, 0) == ""
    assert ui._completion_pct(0, 4) == "0.0"
    with pytest.raises(ValueError, match="Invalid"):
        ui._completion_pct(5, 4)


def test_patient_export_shows_name_and_percentage_without_writing_workbook(monkeypatch, tmp_path):
    workbook = tmp_path / "frozen.xlsm"
    workbook.write_bytes(b"do-not-change")
    destination = tmp_path / "patient.tsv"
    monkeypatch.setattr(ui, "read_patient_history_report", lambda *args, **kwargs: (
        _row("01"), _row("2", scheduled=0, performed=0), _row("old"),
    ))
    monkeypatch.setattr(ui, "read_patient_registry", lambda *args, **kwargs: [
        _patient("1", "ΑΝΝΑ ΠΑΠΑΔΟΠΟΥΛΟΥ"), _patient("2", "ΝΙΚΟΣ"),
    ])
    result = ui.export_patient_history_ui(workbook, destination, mode="overall")
    assert result["rows"] == 3
    lines = destination.read_text(encoding="utf-8-sig").splitlines()
    assert lines[0].split("\t") == list(ui.HEADERS)
    assert lines[1].split("\t")[3] == "ΑΝΝΑ ΠΑΠΑΔΟΠΟΥΛΟΥ"
    assert lines[1].split("\t")[-1] == "75.0"
    assert lines[2].split("\t")[-1] == ""
    assert lines[3].split("\t")[3] == "Άγνωστο όνομα (κωδικός old)"
    assert workbook.read_bytes() == b"do-not-change"


def test_patient_vba_renders_name_not_identifier_and_formats_percentage():
    assert 'ws.Cells(rowIndex, 3).Value2 = fields(3)' in PATIENT_REPORT_VBA
    assert 'If UBound(fields) <> 10' in PATIENT_REPORT_VBA
    assert '"Συμμετοχή %"' in PATIENT_REPORT_VBA
    assert 'ws.Cells(rowIndex, 10).NumberFormat = "0.0%"' in PATIENT_REPORT_VBA
    assert 'ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True' in PATIENT_REPORT_VBA


def test_patient_chart_is_native_idempotent_and_stays_above_detail_table():
    assert 'ws.ChartObjects("ekaPatientOverviewChart").Delete' in PATIENT_REPORT_VBA
    assert 'ws.ChartObjects.Add' in PATIENT_REPORT_VBA
    assert 'series.Values = Array(scheduled, delivered, absent)' in PATIENT_REPORT_VBA
    assert 'series.XValues = Array("Προγραμματισμένες", "Πραγματοποιημένες", "Απουσίες")' in PATIENT_REPORT_VBA
    assert 'ws.Range("A22:J22").Value = Array(' in PATIENT_REPORT_VBA
    assert 'rowIndex = 23' in PATIENT_REPORT_VBA
    assert 'ws.Range("A6").Value2 = "Δεν υπάρχουν οριστικοποιημένες συνεδρίες για γράφημα."' in PATIENT_REPORT_VBA


def test_patient_selection_drilldown_preserves_protected_report():
    assert "Public Sub ShowSelectedPatientChart()" in PATIENT_REPORT_VBA
    assert '"PersonChart", "Γράφημα ασθενή", "ShowSelectedPatientChart"' in PATIENT_REPORT_VBA
    assert 'selectedRow < 23 Or selectedRow > finalRow' in PATIENT_REPORT_VBA
    assert 'BuildPatientOverviewChart ws, CLng(ws.Cells(selectedRow, 5).Value2)' in PATIENT_REPORT_VBA
    assert 'If Not ws Is Nothing Then ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True' in PATIENT_REPORT_VBA
    assert 'ActiveWindow.SplitRow = 22' in PATIENT_REPORT_VBA
