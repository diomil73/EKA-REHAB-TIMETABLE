import inspect
from datetime import date
from types import SimpleNamespace

from rehab_excel.writeback import CellPatch, TextRunPatch
from tools import prepare_daily_input_in_place as module


def test_daily_input_preparation_only_builds_validated_plan():
    source = inspect.getsource(module.prepare_daily_input_payload)

    assert "assert_application_contract(source)" in source
    assert "read_daily_input" in source
    assert "read_daily_replacement_facts" in source
    assert "build_daily_session_states" in source
    assert "build_patient_centric_preview_plan" in source
    assert "compose_outpatient_daily_plan" in source
    assert "build_master_daily_overlay_plan" in source
    assert "compose_master_daily_overlay" in source
    assert "validate_write_plan(plan)" in source
    assert "file_sha256(source)" in source

    assert "apply_write_plan_to_copy" not in source
    assert "shutil.copy" not in source
    assert "Workbooks.Open" not in source


def test_daily_input_effective_state_preserves_active_replacements():
    source = inspect.getsource(module.prepare_daily_input_payload)

    assert "replacements = read_daily_replacement_facts(" in source
    assert "replacements=replacements" in source
    assert '"active_replacements": len(replacements)' in source


def test_daily_input_plan_includes_operational_master_projection():
    source = inspect.getsource(module.prepare_daily_input_payload)

    assert "master_overlay = build_master_daily_overlay_plan(" in source
    assert "plan = compose_master_daily_overlay(plan, master_overlay)" in source
    assert '"master_projection_included": "MASTER_SCHEDULE" in touched_sheets' in source
    assert '"therapist_projection_included": "THERAPIST_DAILY" in touched_sheets' in source


def test_daily_input_payload_serializes_patch_semantics():
    source = inspect.getsource(module._patch_line)

    assert "patch.sheet" in source
    assert "patch.cell" in source
    assert "patch.intent.value" in source
    assert "patch.font_role" in source
    assert "patch.fill_role" in source
    assert "patch.border_role" in source
    assert "patch.wrap_text" in source
    assert "_payload_min_font_size(patch)" in source
    assert "_text_runs_token(patch)" in source


def test_daily_input_payload_serializes_authoritative_change_records_before_patches():
    source = inspect.getsource(module.prepare_daily_input_payload)
    assert "*(_change_line(change, target_date) for change in daily.source_changes)" in source
    assert "*(_patch_line(patch) for patch in plan.patches)" in source
    assert '"ledger_rows": len(daily.source_changes)' in source


def test_change_line_preserves_change_id_subject_reason_and_source_row():
    change = SimpleNamespace(
        change_id="daily-input:2026-10-31:patient:23",
        change_type="patient_absence",
        subject_id="P23",
        display_name="Πετιρόπουλος",
        reason="ΔΕΝ ΠΡΟΣΗΛΘΕ",
        source_row=23,
        all_day=True,
        start_time=None,
    )
    line = module._change_line(change, date(2026, 10, 31))
    assert line.startswith("CHANGE\tdaily-input:2026-10-31:patient:23\t2026-10-31\tpatient_absence\tP23\tΠετιρόπουλος")
    assert "Κανονικό πρόγραμμα" in line
    assert "Απουσία ασθενούς όλη ημέρα" in line
    assert line.endswith("ΔΕΝ ΠΡΟΣΗΛΘΕ\t23")


def test_rich_text_payload_skips_whole_cell_min_font_size():
    patch = CellPatch(
        sheet="THERAPIST_DAILY",
        cell="B2",
        value="A\nB",
        min_font_size=10,
        text_runs=(TextRunPatch(start=1, length=1, font_role="muted"),),
    )
    plain = CellPatch(
        sheet="THERAPIST_DAILY",
        cell="B3",
        value="A",
        min_font_size=10,
    )

    assert module._payload_min_font_size(patch) is None
    assert module._payload_min_font_size(plain) == 10


def test_daily_input_payload_escape_contract():
    escaped = module._escape("A\\B\tC\nD\rE")
    assert escaped == "A\\\\B\\tC\\nD\\rE"
