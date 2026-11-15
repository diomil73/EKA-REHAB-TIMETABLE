"""Minimal native afternoon panel preview; never updates source or overwrite output."""
from __future__ import annotations

import argparse
from pathlib import Path

from rehab_excel.afternoon_minimal_native_installer import (
    AfternoonMinimalInstallError, install_afternoon_minimal_preview,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Install only empty PM technical sheet + read-only patient form panel"
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        output = install_afternoon_minimal_preview(args.source, args.output)
    except (AfternoonMinimalInstallError, OSError, ValueError) as exc:
        print("SAFETY STOP: minimal afternoon panel native proof failed")
        print(f"Reason type: {type(exc).__name__}")
        return 2
    print("MINIMAL AFTERNOON READ-ONLY NATIVE PREVIEW OK")
    print(f"Preview: {output}")
    print("Independent afternoon technical sheet: header-only, veryHidden")
    print("Existing MASTER/PATIENT_PLANNER/PATIENTS: values and MASTER widths unchanged")
    print("Existing patient runtime/form: source-patched, native re-open verified")
    print("No afternoon scheduling save or real patient import enabled")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
