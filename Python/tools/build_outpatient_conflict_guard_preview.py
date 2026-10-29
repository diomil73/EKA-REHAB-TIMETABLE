from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.outpatient_conflict_guard_preview import (  # noqa: E402
    OutpatientConflictGuardPreviewError,
    create_outpatient_conflict_guard_preview,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Install the therapist double-booking popup guard into a safe XLSM copy."
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        report = create_outpatient_conflict_guard_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except OutpatientConflictGuardPreviewError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("OUTPATIENT CONFLICT GUARD PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"Outpatient form present: {report.outpatient_form_present}")
    print(f"Conflict popup present: {report.popup_form_present}")
    print(f"Conflict module present: {report.popup_module_present}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
