"""Synthetic OOXML MASTER structural readback tests, no native Excel."""
from dataclasses import replace
from hashlib import sha256

import pytest
from openpyxl import Workbook, load_workbook

from rehab_excel.daily_input_master_projection_readback import (
    SOURCE_COLUMNS, MasterProjectionSavedReadbackError,
    verify_master_projection_saved_readback,
)
from rehab_excel.master_projection import MasterProjectionRow


def fixture(tmp_path, *, expanded=False, doctor=False):
    path = tmp_path / "fictional-master.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER_SCHEDULE"
    ws["B3"], ws["C3"] = "Θάλαμος", "Ασθενής"
    if expanded:
        ws["K3"], ws["L3"], ws["M3"] = "Ψυχολογοι", "Απογευματινο Προγραμμα", "Κατάσταση"
    else:
        ws["K3"] = "Κατάσταση"
    status = 13 if expanded else 11
    items = (
        MasterProjectionRow(
            5, 2, "P-001", "A01", "FICTIONAL PATIENT 1", False, "active",
            responsible_doctor="DOCTOR X" if doctor else None,
            psychology_display="16:30\nΔΕ, ΠΕ\nSYNTHETIC PROVIDER" if expanded else None,
        ),
        MasterProjectionRow(
            6, 3, "P-002", "B01", "FICTIONAL PATIENT 2", False, "active",
            psychology_display="12:00\\nΤΕ, ΠΑ\\nSYNTHETIC PROVIDER" if expanded else None,
        ),
    )
    for item in items:
        for col, source in SOURCE_COLUMNS.items():
            ws.cell(item.target_row, col).value = f"=PATIENT_PLANNER!{source}{item.planner_row}"
        ws.cell(item.target_row, status).value = f"=PATIENT_PLANNER!E{item.planner_row}"
        if expanded:
            ws.cell(item.target_row, 11).value = item.psychology_display
        if item.responsible_doctor:
            ws.cell(item.target_row, 3).value = f"{item.display_name}\n✚ {item.responsible_doctor}"
    wb.save(path)
    wb.close()
    return path, items


@pytest.mark.parametrize("expanded,doctor", [(False, False), (True, False), (True, True)])
def test_saved_master_exact_linkage_and_no_write(tmp_path, expanded, doctor):
    path, items = fixture(tmp_path, expanded=expanded, doctor=doctor)
    original = path.read_bytes()
    result = verify_master_projection_saved_readback(path, expected_projection=items)
    assert result.projected_rows == 2
    assert result.header_row == 3
    assert result.status_column == (13 if expanded else 11)
    assert result.workbook_sha256 == sha256(original).hexdigest()
    assert path.read_bytes() == original


@pytest.mark.parametrize("corruption", [
    "wrong-formula", "literal-instead-formula", "wrong-status", "doctor",
    "unexpected-separator", "unexpected-tail", "legacy-tail",
    "missing-header", "unsupported-status", "duplicate-header",
    "psychology", "afternoon", "psychology-formula",
])
def test_master_rejects_corrupt_saved_cells(tmp_path, corruption):
    path, items = fixture(tmp_path, expanded=True, doctor=True)
    wb = load_workbook(path, keep_vba=False)
    ws = wb["MASTER_SCHEDULE"]
    if corruption == "wrong-formula":
        ws["F5"] = "=PATIENT_PLANNER!N4"
    elif corruption == "literal-instead-formula":
        ws["F5"] = "a static value"
    elif corruption == "wrong-status":
        ws["M6"] = "=PATIENT_PLANNER!E5"
    elif corruption == "doctor":
        ws["C5"] = "OTHER"
    elif corruption == "unexpected-separator":
        ws["B4"] = "INJECTED"
    elif corruption == "unexpected-tail":
        ws["B7"] = "INJECTED"
    elif corruption == "legacy-tail":
        ws["N5"] = "OLD DATA"
    elif corruption == "missing-header":
        ws["C3"] = "UNKNOWN"
    elif corruption == "unsupported-status":
        ws["M3"] = "UNKNOWN"
    elif corruption == "duplicate-header":
        ws["B2"], ws["C2"] = "Θάλαμος", "Ασθενής"
    elif corruption == "psychology":
        ws["K5"] = "WRONG PSYCHOLOGY"
    elif corruption == "afternoon":
        ws["L5"] = "UNAPPROVED PM DATA"
    elif corruption == "psychology-formula":
        ws["K5"] = "=PATIENT_PLANNER!K2"
    wb.save(path)
    wb.close()
    original = path.read_bytes()
    with pytest.raises(MasterProjectionSavedReadbackError):
        verify_master_projection_saved_readback(path, expected_projection=items)
    assert path.read_bytes() == original


@pytest.mark.parametrize("corruption", [
    "duplicate-target", "nonascending", "duplicate-id", "bad-planner-row",
    "bad-target-row", "not-a-tuple",
])
def test_untrusted_expected_projection_is_rejected(tmp_path, corruption):
    path, items = fixture(tmp_path)
    bad = {
        "duplicate-target": (items[0], replace(items[1], target_row=items[0].target_row)),
        "nonascending": tuple(reversed(items)),
        "duplicate-id": (items[0], replace(items[1], patient_id=items[0].patient_id)),
        "bad-planner-row": (replace(items[0], planner_row=1), items[1]),
        "bad-target-row": (replace(items[0], target_row=1), items[1]),
        "not-a-tuple": list(items),
    }[corruption]
    with pytest.raises(MasterProjectionSavedReadbackError):
        verify_master_projection_saved_readback(path, expected_projection=bad)


def test_missing_sheet_fails_closed(tmp_path):
    path, items = fixture(tmp_path)
    wb = load_workbook(path, keep_vba=False)
    del wb["MASTER_SCHEDULE"]
    wb.create_sheet("NOT_MASTER")
    wb.save(path)
    wb.close()
    with pytest.raises(MasterProjectionSavedReadbackError):
        verify_master_projection_saved_readback(path, expected_projection=items)


def test_no_phase_authorization_or_com_claims():
    import inspect
    import rehab_excel.daily_input_master_projection_readback as module
    code = inspect.getsource(module)
    assert "DailyInputJournalApplyEnabled = True" not in code
    assert "DispatchEx" not in code
    assert "ThisWorkbook.Save" not in code
