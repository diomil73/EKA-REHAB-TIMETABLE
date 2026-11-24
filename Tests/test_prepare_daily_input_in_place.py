import inspect
from datetime import date
from types import SimpleNamespace

from pathlib import Path

from rehab_excel.writeback import CellPatch, TextRunPatch, WriteIntent
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


def test_daily_input_filters_discharge_events_before_materializing_sessions():
    source = inspect.getsource(module.prepare_daily_input_payload)
    assert "read_patient_exits(source, patients=patients)" in source
    assert "base_entries = active_entries_on(" in source
    assert source.index("base_entries = active_entries_on(") < source.index(
        "sessions = materialize_sessions_for_date("
    )
    # Permanent patient exits must never be logged as daily absences.
    assert "patient_exit" not in inspect.getsource(module._change_line)


def _synthetic_preparation_inputs(tmp_path, monkeypatch):
    """One isolated fictional workbook, with real TSV serializer and file I/O."""
    from datetime import time
    from openpyxl import Workbook
    from rehab_core.daily_state import DailySessionState, DailySessionStatus
    from rehab_excel.writeback import WritePlan

    source = tmp_path / "FICTIONAL_MASTER_DAILY.xlsm"
    wb = Workbook()
    sheet = wb.active
    sheet.title = "DAILY_INPUT"
    sheet["B2"] = "24/11/2026"
    wb.create_sheet("MASTER_SCHEDULE")
    wb.save(source)
    wb.close()

    change = SimpleNamespace(
        change_id="FICTIONAL-CHANGE-1", change_type="patient_absence",
        subject_id="FICTIONAL-P1", display_name="FICTIONAL PATIENT",
        reason="FICTIONAL REASON", source_row=8, all_day=True, start_time=None,
    )
    state = DailySessionState(
        session_id="FICTIONAL-BASE-1@2026-11-24",
        patient_id="FICTIONAL-P1", session_date=date(2026, 11, 24),
        status=DailySessionStatus.PATIENT_ABSENT,
        original_therapist_id="FICTIONAL-T1", original_time=time(9, 0),
        effective_therapist_id=None, effective_time=None,
    )
    patch = CellPatch(
        sheet="MASTER_SCHEDULE", cell="D7",
        intent=module.WriteIntent.PRESENTATION, strike_through=True,
        source_tag="master_daily_non_active_overlay",
    )
    original = source.read_bytes()
    monkeypatch.setattr(
        module, "assert_application_contract",
        lambda _: SimpleNamespace(component_names=("FICTIONAL",)),
    )
    monkeypatch.setattr(module, "read_patient_registry", lambda _: ())
    monkeypatch.setattr(
        module, "read_settings",
        lambda _: SimpleNamespace(therapist_names=()),
    )
    monkeypatch.setattr(module, "read_patient_exits", lambda *_args, **_kwargs: ())
    monkeypatch.setattr(module, "read_unified_base_schedule", lambda _: ())
    monkeypatch.setattr(module, "active_entries_on", lambda *_args: ())
    monkeypatch.setattr(module, "materialize_sessions_for_date", lambda *_args: ())
    monkeypatch.setattr(
        module, "read_daily_input",
        lambda *_args, **_kwargs: SimpleNamespace(
            target_date=date(2026, 11, 24),
            absences=(), cancellations=(), source_changes=(change,),
            therapist_rows_used=0, patient_rows_used=1,
        ),
    )
    monkeypatch.setattr(module, "read_daily_replacement_facts", lambda *_args, **_kwargs: ())
    monkeypatch.setattr(module, "build_daily_session_states", lambda *_args, **_kwargs: (state,))
    monkeypatch.setattr(
        module, "build_patient_centric_preview_plan",
        lambda *_args, **_kwargs: SimpleNamespace(
            write_plan=WritePlan(str(source), (patch,)), bindings=(state.session_id,),
        ),
    )
    monkeypatch.setattr(
        module, "compose_outpatient_daily_plan", lambda plan, **_kwargs: plan,
    )
    monkeypatch.setattr(
        module, "build_master_daily_overlay_plan",
        lambda *_args, **_kwargs: SimpleNamespace(write_plan=WritePlan(str(source), ())),
    )
    monkeypatch.setattr(module, "compose_master_daily_overlay", lambda plan, _: plan)
    monkeypatch.setattr(module, "validate_write_plan", lambda _: None)
    return source, original


def test_daily_input_payload_is_exclusively_published_after_full_validation(
    tmp_path, monkeypatch,
):
    source, original = _synthetic_preparation_inputs(tmp_path, monkeypatch)
    output = tmp_path / "FICTIONAL_APPLY_PAYLOAD.tsv"
    result = module.prepare_daily_input_payload(source, output)
    assert result["ok"] is True
    assert result["daily_input_changes"] == 1
    assert result["source_unchanged"] is True
    text = output.read_text(encoding="utf-8-sig")
    assert text.startswith("CHANGE\tFICTIONAL-CHANGE-1")
    assert "\nPATCH\tMASTER_SCHEDULE\tD7\tpresentation" in text
    assert source.read_bytes() == original
    assert not tuple(tmp_path.glob(".diomil_unverified_*"))


def test_daily_input_existing_payload_is_preserved_on_refused_retry(
    tmp_path, monkeypatch,
):
    import pytest
    source, original = _synthetic_preparation_inputs(tmp_path, monkeypatch)
    output = tmp_path / "EXISTING_REPORT.tsv"
    output.write_bytes(b"USER OWNED FILE NEVER DELETE")
    with pytest.raises(ValueError, match="already exists"):
        module.prepare_daily_input_payload(source, output)
    assert output.read_bytes() == b"USER OWNED FILE NEVER DELETE"
    assert source.read_bytes() == original
    assert not tuple(tmp_path.glob(".diomil_unverified_*"))


def test_daily_input_source_hash_drift_does_not_publish_unverified_payload(
    tmp_path, monkeypatch,
):
    import pytest
    from rehab_excel.authoritative_commit import file_sha256 as original_sha

    source, original = _synthetic_preparation_inputs(tmp_path, monkeypatch)
    output = tmp_path / "STALENESS_MUST_NOT_EXIST.tsv"
    hash_calls = []
    def source_sha(path):
        hash_calls.append(path)
        return original_sha(path) if len(hash_calls) == 1 else "0" * 64
    monkeypatch.setattr(module, "file_sha256", source_sha)
    with pytest.raises(module.DailyInputInPlacePreparationError, match="changed"):
        module.prepare_daily_input_payload(source, output)
    assert len(hash_calls) == 2
    assert not output.exists()
    assert source.read_bytes() == original
    assert not tuple(tmp_path.glob(".diomil_unverified_*"))


def test_daily_input_partial_staged_write_is_cleaned_without_publishing(
    tmp_path, monkeypatch,
):
    import pytest
    source, original = _synthetic_preparation_inputs(tmp_path, monkeypatch)
    output = tmp_path / "FAILED_PRIVATE_STAGE.tsv"
    original_write = Path.write_text
    def broken_write(path, data, *args, **kwargs):
        if path.name == "stage.tsv":
            path.write_bytes(b"PRIVATE PARTIAL TSV")
            raise OSError("FICTIONAL INTERRUPTED WRITE")
        return original_write(path, data, *args, **kwargs)
    monkeypatch.setattr(Path, "write_text", broken_write)
    with pytest.raises(OSError, match="INTERRUPTED WRITE"):
        module.prepare_daily_input_payload(source, output)
    assert not output.exists()
    assert source.read_bytes() == original
    assert not tuple(tmp_path.glob(".diomil_unverified_*"))


def test_daily_input_symlink_payload_is_refused_without_modifying_target(
    tmp_path, monkeypatch,
):
    import pytest
    source, original = _synthetic_preparation_inputs(tmp_path, monkeypatch)
    existing = tmp_path / "PRIVATE_KEEP.tsv"
    existing.write_bytes(b"PRIVATE USER OWNED")
    alias = tmp_path / "REDIRECTED_PAYLOAD.tsv"
    try:
        alias.symlink_to(existing)
    except (OSError, NotImplementedError):
        pytest.skip("Symlinks require privilege or platform support")
    with pytest.raises(ValueError, match="already exists"):
        module.prepare_daily_input_payload(source, alias)
    assert existing.read_bytes() == b"PRIVATE USER OWNED"
    assert alias.is_symlink()
    assert source.read_bytes() == original


def test_daily_input_cli_failure_does_not_delete_existing_output_or_log_patient(
    tmp_path, monkeypatch, capsys,
):
    import json
    import pytest
    source, original = _synthetic_preparation_inputs(tmp_path, monkeypatch)
    output = tmp_path / "PRIVATE_EXISTING_OUTPUT.tsv"
    response = tmp_path / "FICTIONAL_RESPONSE.json"
    output.write_bytes(b"PREVIOUS USER WORK PRODUCT")
    monkeypatch.setattr(
        module.sys,
        "argv",
        [
            "prepare_daily_input_in_place.py",
            "--input", str(source), "--payload", str(output),
            "--response", str(response),
        ],
    )
    def fail(*args):
        raise RuntimeError("PRIVATE_PATIENT_NAME_DO_NOT_LOG")
    monkeypatch.setattr(module, "prepare_daily_input_payload", fail)
    assert module.main() == 2
    assert output.read_bytes() == b"PREVIOUS USER WORK PRODUCT"
    assert source.read_bytes() == original
    data = json.loads(response.read_text(encoding="utf-8-sig"))
    assert data["ok"] is False
    assert "PRIVATE_PATIENT_NAME_DO_NOT_LOG" not in data["error"]
    stdout = capsys.readouterr().out
    assert "SAFETY STOP" in stdout
    assert "PRIVATE_PATIENT_NAME_DO_NOT_LOG" not in stdout
    assert str(output) not in stdout
