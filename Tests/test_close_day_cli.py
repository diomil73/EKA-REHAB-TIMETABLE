from pathlib import Path


TOOL_PATH = Path("Python/tools/close_day.py")


def test_close_day_cli_persists_snapshot_then_finalization_event_then_rechecks_contract():
    source = TOOL_PATH.read_text(encoding="utf-8")

    preview = source.index("build_finalized_day_snapshot_preview(")
    build_event = source.index("event = build_day_finalized_event(", preview)
    append_event = source.index("append_day_history_events(output, (event,))", build_event)
    reread_event = source.index("read_day_history_events(", append_event)
    contract = source.index("output_contract = assert_application_contract(output)", reread_event)
    source_hash = source.index("file_sha256(source) != source_sha256_before", contract)

    assert preview < build_event < append_event < reread_event < contract < source_hash
    assert 'event_id=f"day-finalized:{snapshot.service_date.isoformat()}:v{snapshot.version}"' in source
    assert '"source_sha256_before": source_sha256_before' in source
    assert '"unresolved_sessions": snapshot.unresolved_sessions' in source


def test_close_day_cli_deletes_preview_on_any_post_preview_failure():
    source = TOOL_PATH.read_text(encoding="utf-8")

    assert "if output.exists():" in source
    assert "output.unlink()" in source
    assert '"ok": False' in source
    assert "SAFETY STOP:" in source
