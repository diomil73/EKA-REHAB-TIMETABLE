from pathlib import Path


TOOL_PATH = Path("Python/tools/build_finalized_day_snapshot_preview.py")


def test_preview_tool_wires_finalization_before_native_persistence_and_readback():
    source = TOOL_PATH.read_text(encoding="utf-8")

    build = source.index("snapshot = build_finalized_day_snapshot(")
    workload = source.index("therapist_history = build_therapist_day_history_rows(", build)
    copy = source.index("shutil.copy2(source, output)")
    append = source.index("append_finalized_day_snapshot(output, snapshot)")
    reread = source.index("read_finalized_day_snapshots(", append)
    workload_append = source.index("append_therapist_day_history_rows(", reread)
    workload_reread = source.index("read_therapist_day_history_rows(", workload_append)
    source_hash = source.index("source_after = _sha256(source)", workload_reread)

    assert build < workload < copy < append < reread < workload_append < workload_reread < source_hash
    assert "if reread != (snapshot,):" in source
    assert "if history_reread != therapist_history:" in source
    assert "Input workbook changed during preview build" in source


def test_preview_tool_reports_real_application_contract_component_names():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "len(source_contract.component_names)" in source
    assert "len(output_contract.component_names)" in source
    assert ".components" not in source


def test_preview_tool_uses_official_therapist_registry_for_daily_input():
    source = TOOL_PATH.read_text(encoding="utf-8")

    settings = source.index("settings = read_settings(source)")
    daily_input = source.index("daily = read_daily_input(", settings)
    therapist_registry = source.index(
        "therapist_names=settings.therapist_names,",
        daily_input,
    )

    assert settings < daily_input < therapist_registry


def test_preview_tool_excludes_legacy_student_placeholders_from_therapist_history():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "not _is_legacy_student_name(name)" in source
    assert "capacity_timeslots_by_therapist={name: 6 for name in therapist_ids}" in source
    assert "Therapist/day history rows:" in source
