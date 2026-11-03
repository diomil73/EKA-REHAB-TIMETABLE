from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
import shutil
import subprocess
import sys


REPO_ROOT = Path(__file__).resolve().parents[2]
INSTALLERS = (
    "install_unified_day_save_ui.py",
    "install_therapist_history_ui.py",
    "install_patient_history_ui.py",
)
DEFAULT_SOURCE = REPO_ROOT / "Excel" / "previews" / "UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm"


class H7IntegratedPreparationError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def prepare_h7_integrated_ui(source: Path, destination: Path) -> None:
    """Create a distinct disposable XLSM; never mutate or reuse source/smoke files.

    Each subprocess uses its own Excel instance, with persisted read-only
    verification provided by the corresponding individual installer.
    """
    source = source.resolve()
    destination = destination.resolve()
    if not source.is_file() or source.suffix.lower() != ".xlsm":
        raise H7IntegratedPreparationError("Existing XLSM source required")
    if destination.exists():
        raise H7IntegratedPreparationError("Destination already exists: refusing overwrite")
    if destination.suffix.lower() != ".xlsm":
        raise H7IntegratedPreparationError("Destination must end in .xlsm")
    if source == destination:
        raise H7IntegratedPreparationError("Source and destination cannot be the same")

    destination.parent.mkdir(parents=True, exist_ok=True)
    # 'xb' will fail if destination appears between checks.
    with source.open("rb") as original, destination.open("xb") as copied:
        shutil.copyfileobj(original, copied)
    if _sha256(source) != _sha256(destination):
        raise H7IntegratedPreparationError("Fresh copy failed SHA-256 verification")
    print("H7 INTEGRATED COPY VERIFIED", flush=True)

    for installer in INSTALLERS:
        script = REPO_ROOT / "Python" / "tools" / installer
        completed = subprocess.run((sys.executable, str(script), str(destination)), check=False)
        if completed.returncode != 0:
            raise H7IntegratedPreparationError(
                f"{installer} failed with exit {completed.returncode}; "
                "freeze destination, do not retry or repair"
            )
    print("H7 INTEGRATED UI INSTALL OK", flush=True)
    print("Source unchanged; fresh destination installed", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="H6+H4+H7 integrated UI fresh smoke setup")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        prepare_h7_integrated_ui(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
