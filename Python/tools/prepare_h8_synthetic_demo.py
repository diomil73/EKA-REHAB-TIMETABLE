from __future__ import annotations

"""Create one fresh native H8 dashboard and seed deterministic illustrative history."""

import argparse
from pathlib import Path
import subprocess
import sys

PYTHON_ROOT = Path(__file__).resolve().parents[1]
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.prepare_h7_integrated_ui_smoke import DEFAULT_SOURCE


class DemoPreparationError(RuntimeError):
    pass


def prepare_demo(source: Path, output: Path) -> None:
    if not source.is_file():
        raise DemoPreparationError("Trusted source workbook missing")
    if output.exists():
        raise DemoPreparationError("Destination already exists; do not overwrite")
    if output.suffix.lower() != ".xlsm" or "_DEMO_" not in output.name.upper():
        raise DemoPreparationError("Destination filename must contain _DEMO_ and end in .xlsm")
    if source.resolve() == output.resolve():
        raise DemoPreparationError("Source/destination collision")

    tools = Path(__file__).resolve().parent
    builder = tools / "prepare_h8_dashboard_smoke.py"
    seeder = tools / "seed_h8_synthetic_history.py"
    steps = (
        ("H8 fresh dashboard build", (sys.executable, str(builder), "--source", str(source), "--output", str(output))),
        ("Synthetic finalized history seed", (sys.executable, str(seeder), "--workbook", str(output))),
    )
    for label, command in steps:
        finished = subprocess.run(command, check=False)
        if finished.returncode != 0:
            raise DemoPreparationError(
                f"{label} failed with exit {finished.returncode}. Freeze output, no retries."
            )
    print("H8 SYNTHETIC DEMO PREPARATION OK")
    print(f"Disposable workbook: {output.resolve()}")
    print("NO PRODUCTION DATA. Demo attendance is synthetic and not wired to chart 3.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build disposable illustrative H8 Excel history")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        prepare_demo(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
