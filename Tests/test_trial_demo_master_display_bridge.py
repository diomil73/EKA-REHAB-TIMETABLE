"""Synthetic-only tests of existing MASTER treatment display in native DEMO.

No real patient, original workbook, native Excel or new layout is created here.
"""
from datetime import time
from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from rehab_excel import trial_demo_master_display_bridge as bridge
from rehab_excel import trial_demo_native_seed as seed
from rehab_excel.patient_program_master_render_vba import MASTER_RENDER_MODULE_CODE


def _projection():
    return SimpleNamespace(
        weekly_slots=SimpleNamespace(rows=(
            ("DEMO-ID", "ΦΘ", "ΔΕ", time(9, 30), "DEMO PROVIDER", "C2"),
            ("DEMO-ID", "ΦΘ", "ΤΕ", time(9, 30), "DEMO PROVIDER", "C2"),
            ("DEMO-ID", "ΦΘ", "ΠΑ", time(9, 30), "DEMO PROVIDER", "C2"),
            ("DEMO-ID", "ΕΦΑ", "ΤΡ", time(10, 15), None, "P2"),
        )),
        patient_registry=SimpleNamespace(rows=(
            ("DEMO-ID", "A01", "DEMO PERSON", "Ο", "Ενεργός", "Εσωτερικός"),
        )),
    )


def _native_sheet(tmp_path: Path, *, invalid: str | None = None) -> Path:
    path = tmp_path / "DIOMIL_DEMO_DISPLAY_SYNTHETIC.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER_SCHEDULE"
    ws.cell(1, 2).value = "Θάλαμος"
    ws.cell(1, 3).value = "Ασθενής"
    fth = ws.cell(3, 4)
    fth.value = "09:30\nΔε-Τε-Πα\nDEMO PROVIDER"
    fth.alignment = fth.alignment.copy(wrap_text=True)
    efa = ws.cell(3, 10)
    efa.value = "10:15\nΤρ"
    efa.alignment = efa.alignment.copy(wrap_text=True)
    if invalid == "missing_day":
        fth.value = "09:30"
    if invalid == "wrong_provider":
        fth.value = "09:30\nΔε-Τε-Πα\nOTHER"
    if invalid == "not_wrapped":
        fth.alignment = fth.alignment.copy(wrap_text=False)
    wb.save(path)
    wb.close()
    return path


def _stub_row(monkeypatch):
    monkeypatch.setattr(
        bridge, "build_master_projection",
        lambda p: (SimpleNamespace(patient_id="DEMO-ID", target_row=3),),
    )


def test_existing_master_renderer_has_authoritative_column_and_line_contract():
    code = MASTER_RENDER_MODULE_CODE
    for treatment, col in bridge._MASTER_TREATMENT_COLUMNS.items():
        if col == 11:  # conditional psychology column depends on expanded schema
            assert 'WriteTreatmentCell master.Cells(masterRow, 11), planner, plannerRow, "Ψυχ"' in code
            continue
        assert (
            f'WriteTreatmentCell master.Cells(masterRow, {col}), '
            f'planner, plannerRow, "{treatment}"'
        ) in code
    assert 'timeCol = FindHeaderColumn(planner, treatment & "_Ώρα")' in code
    assert 'daysCol = FindHeaderColumn(planner, treatment & "_Ημέρες")' in code
    assert "If Len(result) > 0 Then result = result & vbLf" in code
    assert ".WrapText = True" in code


def test_inpatient_approved_slots_follow_existing_line_order():
    projection = _projection()
    from rehab_excel.trial_demo_planner_mirror import choose_planner_demo_mirrors
    result = choose_planner_demo_mirrors(projection.weekly_slots.rows)
    assert {entry.treatment: bridge._expected_lines(entry) for entry in result} == {
        "ΦΘ": "09:30\nΔε-Τε-Πα\nDEMO PROVIDER",
        "ΕΦΑ": "10:15\nΤρ",
    }


def test_independent_saved_readback_exactly_checks_complete_master_cells(
    tmp_path, monkeypatch,
):
    path = _native_sheet(tmp_path)
    _stub_row(monkeypatch)
    original = path.read_bytes()
    assert bridge.verify_private_demo_master_display(path, _projection()) == 2
    assert path.read_bytes() == original


@pytest.mark.parametrize("invalid", ["missing_day", "wrong_provider", "not_wrapped"])
def test_incorrect_display_fails_closed_instead_of_being_accepted(
    tmp_path, monkeypatch, invalid,
):
    path = _native_sheet(tmp_path, invalid=invalid)
    _stub_row(monkeypatch)
    with pytest.raises(bridge.DemoMasterDisplayError, match="MASTER"):
        bridge.verify_private_demo_master_display(path, _projection())


def test_missing_existing_vba_renderer_blocks_instead_of_being_replaced():
    book = SimpleNamespace(VBProject=SimpleNamespace(VBComponents=lambda _: None))
    with pytest.raises(bridge.DemoMasterDisplayError, match="renderer"):
        bridge._require_existing_macro(book)


def test_stale_renderer_does_not_pass_existing_contract():
    class Code:
        CountOfLines = 2
        def Lines(self, first, total):
            return "Public Sub RefreshMasterPatientProgram(ByVal patientId As String)\nEnd Sub"
    component = SimpleNamespace(CodeModule=Code())
    book = SimpleNamespace(
        VBProject=SimpleNamespace(VBComponents=lambda _: component)
    )
    with pytest.raises(bridge.DemoMasterDisplayError, match="differs"):
        bridge._require_existing_macro(book)


def test_production_seed_reuses_existing_renderer_after_projection_and_style():
    import inspect

    source = inspect.getsource(seed.seed_trial_demo_workbook)
    assert "refresh_private_demo_master_display(stage, projection)" in source
    assert "verify_private_demo_master_display(stage, projection)" in source
    assert source.index("refresh_master_projection_in_place(stage)") < source.index(
        "apply_master_visual_style_in_place(stage)"
    ) < source.index("refresh_private_demo_master_display(stage, projection)") < (
        source.index("verify_private_demo_master_display(stage, projection)")
    )
    assert "MASTER_RENDER_MODULE_CODE" not in source
