from __future__ import annotations

import json
from pathlib import Path

from tools.restore_daily_replacement_facts import _write_response


def test_write_response_is_optional(tmp_path: Path):
    _write_response(None, {"ok": True})


def test_write_response_emits_utf8_json_for_excel_bridge(tmp_path: Path):
    response = tmp_path / "nested" / "restore-response.json"

    _write_response(
        response,
        {
            "ok": True,
            "output_path": r"C:\preview\RESTORED.xlsm",
            "source_sha256_before": "abc123",
            "restored_replacement_facts": 2,
            "active_replacements_after_restore": 3,
            "unresolved_sessions_after_recompute": 2,
            "finalized_after_recompute": False,
            "source_unchanged": True,
            "vba_preserved": True,
        },
    )

    payload = json.loads(response.read_text(encoding="utf-8-sig"))
    assert payload == {
        "ok": True,
        "output_path": r"C:\preview\RESTORED.xlsm",
        "source_sha256_before": "abc123",
        "restored_replacement_facts": 2,
        "active_replacements_after_restore": 3,
        "unresolved_sessions_after_recompute": 2,
        "finalized_after_recompute": False,
        "source_unchanged": True,
        "vba_preserved": True,
    }


def test_restore_cli_exposes_excel_bridge_response_contract():
    source = Path("Python/tools/restore_daily_replacement_facts.py").read_text(encoding="utf-8")

    assert '"--response"' in source
    assert '"output_path"' in source
    assert '"source_sha256_before"' in source
    assert '"restored_replacement_facts"' in source
    assert '"active_replacements_after_restore"' in source
    assert '"unresolved_sessions_after_recompute"' in source
    assert '"finalized_after_recompute"' in source
    assert '"source_unchanged"' in source
    assert '"vba_preserved"' in source


def test_restore_bridge_remains_preview_only_before_authoritative_commit():
    source = Path("Python/tools/restore_daily_replacement_facts.py").read_text(encoding="utf-8")

    assert 'raise AuthoritativeDailyRestoreError("Output must differ from source workbook")' in source
    assert "apply_write_plan_to_copy(" in source
    assert "source_before = _sha256(source)" in source
    assert "source_after = _sha256(source)" in source
