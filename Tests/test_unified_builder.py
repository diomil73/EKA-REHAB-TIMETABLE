from __future__ import annotations

from datetime import date
from pathlib import Path
import sys

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]
TOOLS_ROOT = REPO_ROOT / "Python" / "tools"
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import build_unified_preview as unified  # noqa: E402


def test_unified_builder_runs_stages_in_required_order(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "unified.xlsm"
    source.write_bytes(b"source")

    calls: list[tuple[str, list[str]]] = []

    def fake_run(label: str, args: list[str]) -> None:
        calls.append((label, args))

        if "--output" in args:
            output_index = args.index("--output") + 1
            stage_output = Path(args[output_index])
            stage_output.parent.mkdir(parents=True, exist_ok=True)
            stage_output.write_bytes(b"preview")

    monkeypatch.setattr(unified, "_run_step", fake_run)
    monkeypatch.setattr(unified, "_verify_final_workbook", lambda path: None)
    monkeypatch.setattr(unified, "refresh_master_projection_in_place", lambda path: 0)
    monkeypatch.setattr(unified, "finalize_user_navigation", lambda path: None)
    monkeypatch.setattr(unified, "apply_master_visual_style_in_place", lambda path: None)

    result = unified.build_unified_preview(
        source,
        output,
        target_date=date(2026, 9, 28),
        overwrite=True,
    )

    assert result == output.resolve()
    assert [label for label, _ in calls] == [
        "registration menu + forms",
        "patient registration bridge",
        "patient edit flow",
        "outpatient schedule form",
        "outpatient conflict guard",
        "DAILY_INPUT sheet",
    ]
    assert calls[0][1][1].endswith("build_registration_menu_preview.py")
    assert calls[1][1][1].endswith("build_patient_registration_bridge_preview.py")
    assert calls[2][1][1].endswith("build_patient_edit_flow_preview.py")
    assert calls[3][1][1].endswith("build_outpatient_schedule_form_preview.py")
    assert calls[4][1][1].endswith("build_outpatient_conflict_guard_preview.py")
    assert calls[5][1][1].endswith("preview_daily_input_sheet.py")
    assert "2026-09-28" in calls[5][1]


def test_unified_builder_refuses_source_equal_to_output(tmp_path):
    source = tmp_path / "same.xlsm"
    source.write_bytes(b"source")

    with pytest.raises(unified.UnifiedBuildError, match="different from source"):
        unified.build_unified_preview(
            source,
            source,
            target_date=date(2026, 9, 28),
        )


def test_unified_builder_refuses_existing_output_without_overwrite(tmp_path):
    source = tmp_path / "source.xlsm"
    output = tmp_path / "unified.xlsm"
    source.write_bytes(b"source")
    output.write_bytes(b"existing")

    with pytest.raises(unified.UnifiedBuildError, match="already exists"):
        unified.build_unified_preview(
            source,
            output,
            target_date=date(2026, 9, 28),
            overwrite=False,
        )


def test_unified_builder_requires_xlsm(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "unified.xlsm"
    source.write_bytes(b"source")

    with pytest.raises(unified.UnifiedBuildError, match="both be .xlsm"):
        unified.build_unified_preview(
            source,
            output,
            target_date=date(2026, 9, 28),
        )
