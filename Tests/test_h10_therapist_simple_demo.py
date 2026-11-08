from pathlib import Path

import pytest

from rehab_excel.therapist_history_report_vba import THERAPIST_HISTORY_MODULE_CODE
from tools.prepare_h10_therapist_simple_demo import (
    EVENT_CODE, H10DemoError, REQUIRED_VBA, check_candidate,
)


def test_sticky_header_matches_patient_chart_behavior():
    assert "ActiveWindow.SplitRow = 22" in THERAPIST_HISTORY_MODULE_CODE
    assert 'ws.Range("A23").Select' in THERAPIST_HISTORY_MODULE_CODE
    assert "ActiveWindow.FreezePanes = True" in THERAPIST_HISTORY_MODULE_CODE


def test_selected_therapist_highlight_uses_one_active_row_and_resets_on_rerender():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert 'ws.Range("P1").Value2 = selectedRow' in vba
    assert 'ws.Range("P1").Value2 = 0' in vba
    assert 'Formula1:="=ROW()=$P$1"' in vba
    assert 'RGB(37, 99, 170)' in vba
    assert 'ws.Columns("P").Hidden = True' in vba
    assert 'Target.CountLarge <> 1' in EVENT_CODE
    assert "ShowSelectedTherapistChart" in EVENT_CODE


def test_period_benchmarks_are_separate_and_day_weighted_for_patients():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert "periodKey = CStr(ws.Cells(r, 1).Value2)" in vba
    assert "If selectedRow = 0 Or periodKey = selectedPeriod Then" in vba
    assert "referencePatients / referenceDays" in vba
    assert "referencePlanned / referenceCount" in vba
    assert "referenceActual / referenceCount" in vba
    assert '"Μ.Ο. περιόδου: "' in vba
    assert '"Μ.Ο. προγρ."' in vba
    assert '"Μ.Ο. πραγμ."' in vba


def test_simple_labels_keep_technical_cause_out_of_primary_load_chart():
    vba = THERAPIST_HISTORY_MODULE_CODE
    assert vba.count('"Απουσίες/αναβολές"') == 3
    assert vba.count('"Αναστολές θεραπείας"') == 3
    assert vba.count('"Ασθενείς/ημέρα"') == 3
    assert '"Αποδεσμεύσεις"' not in vba
    assert '"Μη διαθέσιμο ιστορικό απουσιών / εργασίμων"' in vba
    assert "ws.Range(\"E19\").Value2" in vba


def test_h10_requires_strict_new_demo_destination_and_existing_bo(tmp_path):
    bo = tmp_path / "H8_DEMO_BO.xlsm"
    bo.write_bytes(b"sentinel")
    bp = tmp_path / "H10_DEMO_BP.xlsm"
    check_candidate(bo, bp)
    bp.write_bytes(b"already exists")
    with pytest.raises(H10DemoError, match="already exists"):
        check_candidate(bo, bp)
    assert bp.read_bytes() == b"already exists"
    with pytest.raises(H10DemoError, match="BP"):
        check_candidate(bo, tmp_path / "H10_DEMO_BQ.xlsm")
    bn = tmp_path / "H8_DEMO_BN.xlsm"
    bn.write_bytes(b"frozen")
    with pytest.raises(H10DemoError, match="BO"):
        check_candidate(bn, tmp_path / "H10_DEMO_BP2.xlsm")


def test_install_requires_persisted_vba_markers_and_event_module():
    assert REQUIRED_VBA
    assert all(fragment in THERAPIST_HISTORY_MODULE_CODE for fragment in REQUIRED_VBA)
    assert EVENT_CODE.count("Worksheet_SelectionChange") == 1
    assert "ThisWorkbook.Save" in THERAPIST_HISTORY_MODULE_CODE
