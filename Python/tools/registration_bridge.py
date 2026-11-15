from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PYTHON_ROOT = SCRIPT_DIR.parent
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.registration_bridge import RegistrationBridgeError, run_registration_bridge


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run one safe registration request for the Excel/VBA UI bridge."
    )
    parser.add_argument("--request", required=True, type=Path, help="UTF-8 JSON request file")
    parser.add_argument("--response", required=True, type=Path, help="UTF-8 JSON response file")
    return parser


def _write_response(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    args = _parser().parse_args()
    try:
        payload = json.loads(args.request.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict):
            raise RegistrationBridgeError("request JSON root must be an object")
        response = run_registration_bridge(payload)
    except (OSError, json.JSONDecodeError, RegistrationBridgeError) as exc:
        _write_response(
            args.response,
            {
                "ok": False,
                "error": str(exc),
            },
        )
        print(f"REGISTRATION BRIDGE STOP: {exc}")
        return 2

    _write_response(args.response, response)
    print("REGISTRATION BRIDGE OK")
    print(f"Kind: {response['kind']}")
    print(f"Subject: {response['subject_key']}")
    print(f"Preview: {response['output_path']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
