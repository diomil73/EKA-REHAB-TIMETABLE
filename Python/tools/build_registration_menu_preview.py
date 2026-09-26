from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.registration_menu_vba import (  # noqa: E402
    DEFAULT_PREVIEW_FILENAME,
    RegistrationMenuVbaError,
    create_registration_menu_preview,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build a safe .xlsm preview containing the central registration UserForm."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / DEFAULT_PREVIEW_FILENAME,
    )
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        report = create_registration_menu_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except RegistrationMenuVbaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("REGISTRATION MENU PREVIEW OK")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA project present: {report.vba_present}")
    print(f"Menu module present: {report.module_present}")
    print(f"Menu form present: {report.form_present}")
    print(f"Open Excel and run macro: {report.macro_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
