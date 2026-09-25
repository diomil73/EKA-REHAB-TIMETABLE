from __future__ import annotations

import argparse
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.student_registry import STUDENT_REGISTRY_HEADERS  # noqa: E402
from rehab_excel.student_registry_schema import (  # noqa: E402
    StudentRegistrySchemaError,
    create_student_registry_schema_preview,
)


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe preview copy with the authoritative STUDENTS registry schema."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "STUDENT_REGISTRY_SCHEMA_PREVIEW.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        report = create_student_registry_schema_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except StudentRegistrySchemaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    output = Path(report.output_path)
    vba_present = _has_vba(output)

    print("STUDENT REGISTRY SCHEMA PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Sheet created: {report.sheet_created}")
    print(f"Headers: {' | '.join(STUDENT_REGISTRY_HEADERS)}")
    print(f"Student rows: {report.student_count}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Schema verified: {report.schema_verified}")
    print(f"VBA project present: {vba_present}")
    print("NEXT: open only STUDENT_REGISTRY_SCHEMA_PREVIEW.xlsm and inspect STUDENTS.")
    return (
        0
        if report.source_unchanged and report.schema_verified and vba_present
        else 3
    )


if __name__ == "__main__":
    raise SystemExit(main())
