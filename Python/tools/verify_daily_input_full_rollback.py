"""Read-only two-stage whole-workbook rollback audit, never an Apply approval.

Requires a retained PREPARED snapshot and a source-bound SHA-256 receipt
created outside this tool. The archive-level comparison deliberately rejects
all changes outside the journal worksheet, even if caused by a benign Excel
save. After the archive audit, native Excel is opened in a NEW, macro-disabled
read-only instance to verify the protected journal rows again.

No workbook is ever created, edited, restored, or promoted by this CLI.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_full_rollback_audit import (  # noqa: E402
    FullRollbackAuditError,
    _read_journal,
    audit_full_workbook_rollback,
    audit_restored_business_state_before_rollback,
)
from tools.verify_daily_input_journal_prepared import (  # noqa: E402
    verify_native_excel_readonly,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Strict native Excel read-only before/after rollback proof."
    )
    parser.add_argument("--snapshot", type=Path, required=True)
    parser.add_argument("--restored", type=Path, required=True)
    parser.add_argument("--snapshot-sha256", type=str, required=True)
    parser.add_argument("--stage", choices=["before-final-row", "final"], required=True)
    args = parser.parse_args(argv)
    try:
        if args.stage == "before-final-row":
            result = audit_restored_business_state_before_rollback(
                args.snapshot,
                args.restored,
                expected_snapshot_sha256=args.snapshot_sha256,
            )
        else:
            result = audit_full_workbook_rollback(
                args.snapshot,
                args.restored,
                expected_snapshot_sha256=args.snapshot_sha256,
            )
        expected_rows = _read_journal(args.restored)
        verify_native_excel_readonly(args.restored, expected_rows)
        if (file_sha256(args.snapshot) != result.snapshot_sha256
                or file_sha256(args.restored) != result.restored_sha256):
            raise FullRollbackAuditError("Source or restored workbook changed during native readback")
    except Exception:
        # No sensitive source paths, patient data or workbook contents in stdout.
        print("SAFETY STOP: FULL-WORKBOOK ROLLBACK AUDIT FAILED")
        return 2

    print("DAILY_INPUT FULL-WORKBOOK ROLLBACK AUDIT GREEN")
    print("Audit stage: " + args.stage)
    print("Nonjournal workbook members matched: " + str(result.checked_nonjournal_parts))
    print("Journal revisions, before: " + str(result.original_journal_revisions))
    print("Journal revisions, after: " + str(result.restored_journal_revisions))
    print("Native Excel read-only journal verifier: GREEN")
    print("Workbook files unchanged by audit: YES")
    print("Operational Apply/Restore executed: NO")
    print("DIAGNOSTIC ONLY: not a durable receipt or clinical approval")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
