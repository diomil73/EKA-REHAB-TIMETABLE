from __future__ import annotations

import importlib.util
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "Python" / "tools" / "registration_bridge_cli.py"


def _load_cli_module():
    spec = importlib.util.spec_from_file_location("registration_bridge_cli", MODULE_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_read_request_accepts_utf8_bom_from_vba_adodb_stream(tmp_path):
    module = _load_cli_module()
    request = tmp_path / "request.json"
    request.write_text(
        '{"action":"new_patient","values":{"display_name":"ΔΟΚΙΜΗ"}}',
        encoding="utf-8-sig",
    )

    payload = module._read_request(request)

    assert payload["action"] == "new_patient"
    assert payload["values"]["display_name"] == "ΔΟΚΙΜΗ"


def test_read_request_still_accepts_plain_utf8(tmp_path):
    module = _load_cli_module()
    request = tmp_path / "request.json"
    request.write_text('{"ok":true}', encoding="utf-8")

    assert module._read_request(request) == {"ok": True}
