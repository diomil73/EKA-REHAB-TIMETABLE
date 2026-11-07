from __future__ import annotations

"""Build one new H8 XLSM from trusted AZ; never edit frozen H7 evidence."""

import argparse
from pathlib import Path
import subprocess
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.prepare_h7_integrated_ui_smoke import (
    DEFAULT_SOURCE,
    H7IntegratedPreparationError,
    REPO_ROOT,
    prepare_h7_integrated_ui,
)


class H8DashboardPreparationError(RuntimeError):
    pass


def prepare_h8_dashboard_smoke(source: Path, destination: Path) -> None:
    """Run H7 installers then H8 guarded VBA patch on a fresh one-shot copy.

    The H7 helper verifies source SHA256 and refuses an existing destination.
    If any stage fails, freeze the destination without retry or repair.
    """
    try:
        prepare_h7_integrated_ui(source, destination)
    except H7IntegratedPreparationError as exc:
        raise H8DashboardPreparationError(str(exc)) from exc

    installer = REPO_ROOT / "Python" / "tools" / "install_h8_single_message_refresh.py"
    completed = subprocess.run((sys.executable, str(installer), str(destination.resolve())), check=False)
    if completed.returncode != 0:
        raise H8DashboardPreparationError(
            f"H8 refresh installer exit {completed.returncode}; freeze destination"
        )
    print("H8 DASHBOARD PREPARATION OK", flush=True)
    print("Fresh destination only; no frozen evidence modified", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Fresh H8 dashboard native integration smoke")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        prepare_h8_dashboard_smoke(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
