from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.master_projection_writer import (  # noqa: E402
    MasterProjectionWriteError,
    create_master_projection_preview,
)


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a safe MASTER_SCHEDULE projection preview."
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "MASTER_PROJECTION_SMOKE.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()

    try:
        report = create_master_projection_preview(
            source,
            output,
            overwrite=args.overwrite,
        )
    except MasterProjectionWriteError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    vba_ok = _has_vba(output)

    print("MASTER PROJECTION PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Projected rows: {report.projected_rows}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Read-back verified: {report.verified_in_output}")
    print(f"VBA preserved: {vba_ok}")
    print("NEXT: open only MASTER_PROJECTION_SMOKE.xlsm and inspect MASTER_SCHEDULE.")

    return 0 if (
        report.source_unchanged
        and report.verified_in_output
        and vba_ok
    ) else 3


if __name__ == "__main__":
    raise SystemExit(main())
