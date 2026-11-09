from __future__ import annotations

"""Fail-closed review of locally collected pilot acceptance evidence.

This validates an operator's evidence manifest, NOT Excel behavior.
No clinical/patient data, workbook modification, installation, or release.
"""

import argparse
from datetime import date
import json
from pathlib import Path
import re

REQUIRED_GATES = (
    "intake_schema_and_registry",
    "history_origin_isolation",
    "source_copy_and_hash",
    "therapist_capacity_and_slot_collisions",
    "patient_identity_and_session_counts",
    "patient_absence_and_cancellation",
    "therapist_absence_and_replacement",
    "double_booking_and_approval",
    "daily_finalization_and_reopen",
    "amendment_restore_and_rollback",
    "vba_protection_and_ui_navigation",
    "report_totals_and_calendar_scope",
    "primary_workbook_write_isolation",
)
SHA256_PATTERN = re.compile(r"^[a-fA-F0-9]{64}$")
EVIDENCE_PATTERN = re.compile(r"^[A-Za-z0-9_./:-]{6,160}$")
ACCEPTED_STATES = {"pass", "fail", "pending", "blocked"}


class PilotEvidenceError(ValueError):
    pass


def new_manifest() -> dict:
    return {
        "schema_version": 1,
        "candidate_sha256": "",
        "reference_sha256": "",
        "pilot_start_date": "",
        "native_excel_checked": False,
        "gates": [
            {"id": gate, "status": "pending", "evidence_ref": ""}
            for gate in REQUIRED_GATES
        ],
    }


def validate_manifest(document: object) -> tuple[str, ...]:
    if not isinstance(document, dict):
        raise PilotEvidenceError("Manifest root must be an object")
    if set(document) != {
        "schema_version", "candidate_sha256", "reference_sha256",
        "pilot_start_date", "native_excel_checked", "gates",
    }:
        raise PilotEvidenceError("Manifest fields differ from approved contract")
    if document["schema_version"] != 1 or type(document["native_excel_checked"]) is not bool:
        raise PilotEvidenceError("Invalid manifest version or native Excel evidence type")
    if not isinstance(document["gates"], list):
        raise PilotEvidenceError("gates must be a list")
    for field in ("candidate_sha256", "reference_sha256"):
        value = document[field]
        if not isinstance(value, str):
            raise PilotEvidenceError(f"Invalid {field} type")
        if value and not SHA256_PATTERN.fullmatch(value):
            raise PilotEvidenceError(f"Invalid {field} format")
    value = document["pilot_start_date"]
    if not isinstance(value, str):
        raise PilotEvidenceError("Invalid pilot_start_date type")
    if value:
        try:
            if date.fromisoformat(value).isoformat() != value:
                raise ValueError("Not canonical")
        except ValueError as exc:
            raise PilotEvidenceError("Invalid pilot_start_date") from exc
    ids = []
    pending = []
    for entry in document["gates"]:
        if not isinstance(entry, dict) or set(entry) != {"id", "status", "evidence_ref"}:
            raise PilotEvidenceError("Invalid gate structure")
        gate_id, state, evidence = entry["id"], entry["status"], entry["evidence_ref"]
        if not isinstance(gate_id, str) or gate_id not in REQUIRED_GATES:
            raise PilotEvidenceError("Unknown gate")
        if not isinstance(state, str) or state not in ACCEPTED_STATES:
            raise PilotEvidenceError("Invalid gate state")
        if not isinstance(evidence, str):
            raise PilotEvidenceError("Invalid evidence reference")
        if evidence and not EVIDENCE_PATTERN.fullmatch(evidence):
            raise PilotEvidenceError("Evidence reference must be a short non-personal token")
        if state == "pass" and not evidence:
            pending.append(gate_id)
        if state != "pass":
            pending.append(gate_id)
        ids.append(gate_id)
    if len(ids) != len(set(ids)) or set(ids) != set(REQUIRED_GATES):
        raise PilotEvidenceError("Missing or duplicate required acceptance gates")
    if not document["candidate_sha256"] or not document["reference_sha256"]:
        pending.append("workbook_hash_evidence")
    if document["candidate_sha256"] and document["reference_sha256"] and (
        document["candidate_sha256"].lower() == document["reference_sha256"].lower()
    ):
        pending.append("separate_candidate_required")
    if not value:
        pending.append("pilot_start_date")
    if not document["native_excel_checked"]:
        pending.append("native_excel_checked")
    return tuple(dict.fromkeys(pending))


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate local pilot acceptance checklist evidence")
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--create-template", type=Path)
    action.add_argument("--manifest", type=Path)
    args = parser.parse_args()
    if args.create_template:
        destination = args.create_template
        if destination.suffix.lower() != ".json":
            print("SAFETY STOP: Template must be .json")
            return 2
        try:
            with destination.open("x", encoding="utf-8") as stream:
                json.dump(new_manifest(), stream, indent=2, ensure_ascii=False)
                stream.write("\n")
        except (FileExistsError, OSError) as exc:
            print(f"SAFETY STOP: {type(exc).__name__}; never overwrite manifest")
            return 2
        print("PILOT EVIDENCE TEMPLATE CREATED (PENDING, NOT APPROVED)")
        return 0
    try:
        path = args.manifest
        if not path.is_file() or path.suffix.lower() != ".json":
            raise PilotEvidenceError("Existing .json evidence manifest required")
        with path.open("r", encoding="utf-8") as stream:
            evidence = json.load(stream)
        pending = validate_manifest(evidence)
    except (OSError, UnicodeError, json.JSONDecodeError, PilotEvidenceError) as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
    print("PILOT EVIDENCE MANIFEST STRUCTURE OK")
    print(f"Required gates: {len(REQUIRED_GATES)}")
    print(f"Unresolved gate/metadata conditions: {len(pending)}")
    for code in pending:
        print(f"NOT READY: {code}")
    if pending:
        print("PILOT ACCEPTANCE NOT ESTABLISHED")
        return 3
    print("EVIDENCE CHECKLIST COMPLETE; INDEPENDENT NATIVE REVIEW STILL REQUIRED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
