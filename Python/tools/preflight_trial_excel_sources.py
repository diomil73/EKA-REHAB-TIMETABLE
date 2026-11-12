"""Read-only, fail-closed combined preflight for a *separate* native Excel pilot.

Run on institution-controlled storage. No file names, patient names, room
identifiers, provider names or PatientIDs are printed. This command NEVER
opens a writable workbook, modifies the CC, or imports data into MASTER.

A clean preflight still requires separate authorized Windows Excel/COM proof.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_excel.trial_admission_gate import evaluate_trial_admission  # noqa: E402
from tools.audit_pilot_source_readiness import inspect_source  # noqa: E402


def _digest(path: Path) -> str:
    value = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def inspect_native_pilot_inputs(*, review: Path, cc: Path) -> dict[str, object]:
    """Return non-identifying preflight counts; always verify source hashes."""
    review, cc = review.resolve(), cc.resolve()
    if review == cc or review.suffix.lower() != ".xlsx" or cc.suffix.lower() != ".xlsm":
        raise ValueError("Two distinct review.xlsx and cc.xlsm inputs are required")
    if not review.is_file() or not cc.is_file():
        raise ValueError("One or more pilot inputs are missing")
    before = (_digest(review), _digest(cc))
    try:
        gate = evaluate_trial_admission(review)
        master = inspect_source(cc)
    finally:
        if before != (_digest(review), _digest(cc)):
            raise RuntimeError("SAFETY STOP: an input changed during read-only preflight")
    # Do not include raw gate blockers, which may contain row provenance.
    gate_ready = gate.ready
    master_ready = not master["issues"]
    return {
        "review_patients": gate.candidate_patients,
        "verified_patient_ids": gate.verified_patient_ids,
        "proposed_weekly_slots": gate.weekly_proposals,
        "unadmitted_weekly_slots": gate.unadmitted_weekly_rows,
        "source_holds": gate.unresolved_source_rows,
        "source_overlaps": gate.overlap_rows,
        "review_blocker_count": len(gate.blockers),
        "legacy_operational_sheets": master["operational_sheets_found"],
        "legacy_history_rows": master["history_rows"],
        "legacy_issue_codes": master["issues"],
        "review_gate_ready": gate_ready,
        "cc_intake_ready": master_ready,
        "native_trial_permitted": gate_ready and master_ready,
        "inputs_unchanged": True,
    }


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="READ-ONLY two-workbook Diomil pilot gate")
    p.add_argument("--review", required=True, type=Path)
    p.add_argument("--cc", required=True, type=Path)
    args = p.parse_args(argv)
    try:
        status = inspect_native_pilot_inputs(review=args.review, cc=args.cc)
    except Exception as exc:
        # Never log source names/ids/paths from raw workbook exceptions.
        print(f"SAFETY STOP: Preflight could not be completed ({type(exc).__name__}).")
        return 2
    print("DIOMIL READ-ONLY NATIVE TRIAL INPUT CHECK")
    for key in (
        "review_patients", "verified_patient_ids", "proposed_weekly_slots",
        "unadmitted_weekly_slots", "source_holds", "source_overlaps",
        "review_blocker_count", "legacy_operational_sheets", "legacy_history_rows",
    ):
        print(f"{key}: {status[key]}")
    for code in status["legacy_issue_codes"]:
        print(f"CC REVIEW REQUIRED: {code}")
    print("Both source files unchanged: YES")
    if not status["native_trial_permitted"]:
        print("SAFETY STOP: No native Excel import or release is authorized.")
        return 2
    print("SOURCE PREVIEW PASS ONLY: independent native COM trial remains mandatory.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
