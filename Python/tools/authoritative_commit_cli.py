from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.authoritative_commit_bridge import (  # noqa: E402
    AuthoritativeCommitBridgeError,
    run_authoritative_commit_bridge,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Promote one verified XLSM preview to the authoritative workbook."
    )
    parser.add_argument("--request", type=Path, required=True)
    parser.add_argument("--response", type=Path, required=True)
    return parser


def _write_response(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> int:
    args = _build_parser().parse_args()
    try:
        payload = json.loads(args.request.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict):
            raise AuthoritativeCommitBridgeError("request JSON must be an object")
        response = run_authoritative_commit_bridge(payload)
    except (OSError, json.JSONDecodeError, AuthoritativeCommitBridgeError) as exc:
        _write_response(args.response, {"ok": False, "error": str(exc)})
        return 2

    _write_response(args.response, response)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
