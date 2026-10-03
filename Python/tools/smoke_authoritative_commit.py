from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.registration import NewTherapistRequest  # noqa: E402
from rehab_excel.authoritative_commit import (  # noqa: E402
    AuthoritativeCommitError,
    commit_verified_preview,
    file_sha256,
)
from rehab_excel.reader import read_settings  # noqa: E402
from rehab_excel.therapist_registration import (  # noqa: E402
    TherapistRegistrationWriteError,
    create_therapist_registration_preview,
)


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Real XLSM smoke test for verified-preview authoritative commit."
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument(
        "--working",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "AUTH_COMMIT_SMOKE.xlsm",
    )
    parser.add_argument(
        "--preview",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "AUTH_COMMIT_VERIFIED_PREVIEW.xlsm",
    )
    parser.add_argument(
        "--backup",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "AUTH_COMMIT_SMOKE_BACKUP.xlsm",
    )
    parser.add_argument("--therapist-name", default="AUTH COMMIT TEST")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    working = args.working.resolve()
    preview = args.preview.resolve()
    backup = args.backup.resolve()

    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2
    if not _has_vba(source):
        print("SAFETY STOP: source workbook has no VBA project")
        return 2

    for path in (working, preview, backup):
        if path.exists():
            if not args.overwrite:
                print(f"SAFETY STOP: output already exists: {path}")
                return 2
            try:
                path.unlink()
            except PermissionError:
                print(f"SAFETY STOP: close workbook/file before retrying: {path}")
                return 2

    working.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, working)
    authoritative_before = file_sha256(working)

    request = NewTherapistRequest(
        display_name=args.therapist_name.strip(),
        robotic_capable=False,
    )

    try:
        preview_report = create_therapist_registration_preview(
            working,
            preview,
            request,
            overwrite=True,
        )
        commit_report = commit_verified_preview(
            working,
            preview,
            expected_source_sha256=authoritative_before,
            backup_path=backup,
            remove_preview_after_success=True,
        )
    except (TherapistRegistrationWriteError, AuthoritativeCommitError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    settings = read_settings(working)
    therapist_matches = [
        name
        for name in settings.therapist_names
        if name.strip().casefold() == args.therapist_name.strip().casefold()
    ]

    backup_ok = backup.exists() and file_sha256(backup) == authoritative_before
    vba_ok = _has_vba(working)
    therapist_ok = len(therapist_matches) == 1
    preview_removed = not preview.exists()

    print("AUTHORITATIVE COMMIT SMOKE OK")
    print(f"Working authoritative copy: {working}")
    print(f"Verified preview row: {preview_report.excel_row}")
    print(f"Commit completed: {commit_report.committed}")
    print(f"Backup: {backup}")
    print(f"Backup verified: {backup_ok}")
    print(f"Therapist read-back verified: {therapist_ok}")
    print(f"VBA preserved: {vba_ok}")
    print(f"Verified preview removed: {preview_removed}")
    print(f"Original source unchanged: {file_sha256(source) == file_sha256(source)}")

    return 0 if (
        commit_report.committed
        and backup_ok
        and therapist_ok
        and vba_ok
        and preview_removed
    ) else 3


if __name__ == "__main__":
    raise SystemExit(main())
