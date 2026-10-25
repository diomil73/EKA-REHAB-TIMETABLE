from pathlib import Path


TOOL_PATH = Path("Python/tools/build_finalized_day_snapshot_preview.py")


def test_preview_tool_wires_finalization_before_native_persistence_and_readback():
    source = TOOL_PATH.read_text(encoding="utf-8")

    build = source.index("snapshot = build_finalized_day_snapshot(")
    copy = source.index("shutil.copy2(source, output)")
    append = source.index("append_finalized_day_snapshot(output, snapshot)")
    reread = source.index("read_finalized_day_snapshots(", append)
    source_hash = source.index("source_after = _sha256(source)", reread)

    assert build < copy < append < reread < source_hash
    assert "if reread != (snapshot,):" in source
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
