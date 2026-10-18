from __future__ import annotations

import argparse
from hashlib import sha256
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
from rehab_excel.outpatient_schedule_edit_vba import (  # noqa: E402
    EDIT_MACRO_NAME,
    install_outpatient_schedule_edit_support_in_workbook,
)
from rehab_excel.outpatient_schedule_form_vba import (  # noqa: E402
    MACRO_NAME,
    OutpatientScheduleFormVbaError,
    create_outpatient_schedule_form_preview,
)


class OutpatientRecurringSmokeBuildError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _safe_unlink(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def build_outpatient_recurring_smoke(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> Path:
    """Build one fresh smoke workbook with create, edit, suggestions and conflict UI.

    The trusted source is never opened for mutation. A short-lived stage copy gets
    the schedule form and edit support; the final copy is then produced while the
    conflict guard is installed. The stage is removed on both success and failure.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise OutpatientRecurringSmokeBuildError(f"Source workbook not found: {source}")
    if source == output:
        raise OutpatientRecurringSmokeBuildError("Output must differ from source")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise OutpatientRecurringSmokeBuildError("Source and output must both be .xlsm")
    if output.exists() and not overwrite:
        raise OutpatientRecurringSmokeBuildError(f"Output already exists: {output}")

    output.parent.mkdir(parents=True, exist_ok=True)
    stage = output.with_name(f".{output.stem}.outpatient-form-stage{output.suffix}")
    _safe_unlink(stage)
    source_before = _sha256(source)

    try:
        create_outpatient_schedule_form_preview(source, stage, overwrite=True)
        install_outpatient_schedule_edit_support_in_workbook(stage)
        create_outpatient_conflict_guard_preview(stage, output, overwrite=overwrite)
    except (
        OutpatientScheduleFormVbaError,
        OutpatientConflictGuardPreviewError,
        RuntimeError,
        OSError,
    ) as exc:
        _safe_unlink(output)
        raise OutpatientRecurringSmokeBuildError(str(exc)) from exc
    finally:
        _safe_unlink(stage)

    if _sha256(source) != source_before:
        _safe_unlink(output)
        raise OutpatientRecurringSmokeBuildError(
            "Source workbook changed during outpatient recurring smoke build"
        )
    if not output.exists():
        raise OutpatientRecurringSmokeBuildError("Smoke workbook was not produced")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Build one fresh XLSM smoke copy with outpatient recurring create/edit, "
            "suggestions and conflict popup support."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        output = build_outpatient_recurring_smoke(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except OutpatientRecurringSmokeBuildError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("OUTPATIENT RECURRING SMOKE BUILD OK")
    print(f"Source: {args.source.resolve()}")
    print(f"Smoke: {output}")
    print(f"Create macro: {MACRO_NAME}")
    print(f"Edit macro: {EDIT_MACRO_NAME}")
    print("Source unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
