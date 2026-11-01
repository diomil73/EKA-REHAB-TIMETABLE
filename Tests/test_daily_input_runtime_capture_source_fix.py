from __future__ import annotations

import importlib.util
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "Python" / "tools" / "apply_daily_input_runtime_capture_source_fix.py"


def _load_module():
    spec = importlib.util.spec_from_file_location("runtime_capture_fix", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_runtime_capture_migrator_preserves_exact_failure_payload_and_response(tmp_path):
    module = _load_module()
    target = tmp_path / "daily_input_action_vba.py"
    target.write_text(
        "prefix\n"
        "    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo CleanUp\n"
        "middle\n"
        "Private Function RefreshOperationalViews() As Boolean\n"
        "suffix\n",
        encoding="utf-8",
    )
    module.TARGET = target

    assert module.main() == 0

    migrated = target.read_text(encoding="utf-8")
    assert "If Not ApplyValidatedDailyInputPayload(payloadPath) Then\n" in migrated
    assert "PreserveDailyInputFailureArtifacts payloadPath, responsePath" in migrated
    assert 'payloadCopy = Environ$("TEMP") & "\\eka_daily_input_last_payload.tsv"' in migrated
    assert 'responseCopy = Environ$("TEMP") & "\\eka_daily_input_last_response.json"' in migrated
    assert "FileCopy payloadPath, payloadCopy" in migrated
    assert "FileCopy responsePath, responseCopy" in migrated
    assert migrated.index("Private Sub PreserveDailyInputFailureArtifacts") < migrated.index(
        "Private Function RefreshOperationalViews() As Boolean"
    )


def test_runtime_capture_migrator_fails_closed_if_marker_is_not_unique(tmp_path):
    module = _load_module()
    target = tmp_path / "daily_input_action_vba.py"
    marker = "    If Not ApplyValidatedDailyInputPayload(payloadPath) Then GoTo CleanUp\n"
    target.write_text(
        marker + marker + "Private Function RefreshOperationalViews() As Boolean\n",
        encoding="utf-8",
    )
    module.TARGET = target

    try:
        module.main()
    except SystemExit as exc:
        assert "marker count=2, expected=1" in str(exc)
    else:
        raise AssertionError("migrator should fail closed when the source marker is ambiguous")
