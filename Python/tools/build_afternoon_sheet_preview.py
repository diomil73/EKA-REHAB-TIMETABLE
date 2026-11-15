"""Explicit native Excel preview builder; never overwrites or imports patients."""
from __future__ import annotations

import argparse
from pathlib import Path

from rehab_excel.afternoon_sheet_preview import (
    AfternoonPreviewError, build_afternoon_sheet_preview,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a fresh isolated Diomil afternoon-sheet test copy"
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    try:
        build_afternoon_sheet_preview(args.source, args.output)
    except (AfternoonPreviewError, OSError) as exc:
        print(f"SAFETY STOP: {type(exc).__name__}")
        return 2
    print("AFTERNOON NATIVE SHEET PREVIEW OK")
    print("New sheet: header-only, veryHidden")
    print("MASTER column widths: unchanged")
    print("Original workbook: read-only source, hash-verified")
    print("This is NOT a working afternoon patient UI or patient migration.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
