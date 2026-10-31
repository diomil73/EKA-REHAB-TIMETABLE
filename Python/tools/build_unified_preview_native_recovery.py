from __future__ import annotations

import argparse
from datetime import date
from hashlib import sha256
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.application_contract import (  # noqa: E402
    ApplicationContractError,
    assert_application_contract,
)
from tools.build_unified_preview import (  # noqa: E402
    UnifiedBuildError,
    build_unified_preview,
)
from tools.patch_native_restore_recovery import (  # noqa: E402
    NativeRestoreRecoveryPatchError,
    patch_workbook,
)


class NativeRecoveryBuildError(RuntimeError):
    pass


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def build_native_recovery_preview(
    source: Path,
    output: Path,
    *,
    target_date: date,
    overwrite: bool = False,
) -> Path:
    source = source.resolve()
    output = output.resolve()
    source_before = _sha256(source)

    try:
        built = build_unified_preview(
            source,
            output,
            target_date=target_date,
            overwrite=overwrite,
        )
        patch_workbook(built)
        assert_application_contract(built)
    except (UnifiedBuildError, NativeRestoreRecoveryPatchError, ApplicationContractError) as exc:
        output.unlink(missing_ok=True)
        raise NativeRecoveryBuildError(str(exc)) from exc

    if _sha256(source) != source_before:
        output.unlink(missing_ok=True)
        raise NativeRecoveryBuildError("Original source workbook changed during native recovery build")
    return built


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Build the unified preview and apply the bounded native restore recovery patch."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        output = build_native_recovery_preview(
            args.source,
            args.output,
            target_date=args.date,
            overwrite=args.overwrite,
        )
    except NativeRecoveryBuildError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    contract = assert_application_contract(output)
    print("NATIVE RECOVERY UNIFIED PREVIEW OK")
    print(f"Source: {args.source.resolve()}")
    print(f"Preview: {output}")
    print(f"DAILY_INPUT date: {args.date.isoformat()}")
    print("Source unchanged: True")
    print(f"Application contract: {len(contract.required_components)}")
    print("Native M recovery patch: verified")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
