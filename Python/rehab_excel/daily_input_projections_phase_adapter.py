"""Non-authorizing PROJECTIONS_REOPEN phase-report adapter.

Runs the existing independent native source-capture, INPUT_AND_LEDGER,
MASTER and complete THERAPIST_DAILY readbacks against one exact saved XLSM.
Publishes only a canonical report; it does not authenticate the verifier,
append receipts/journal rows, or authorize Apply/Restore.
"""
from __future__ import annotations

from datetime import date, datetime, timezone
from pathlib import Path
import os

from .authoritative_commit import file_sha256
from .daily_input_operation import DailyInputOperation, DailyInputOperationPhase
from .daily_input_phase_evidence import (
    EvidenceKind, PhaseEvidenceContractError, required_evidence_for_transition,
)
from .daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS, DailyInputPhaseVerifierReport,
    PhaseVerifierContractError, create_phase_verifier_report,
)
from .daily_input_projection_preflight import DailyInputProjectionPreflight
from .daily_input_projection_preflight_native import (
    DailyInputProjectionPreflightNativeError,
    DailyInputProjectionPreflightNativeReadback,
    verify_daily_input_projection_preflight_native_readonly,
)


class ProjectionsPhaseReportAdapterError(RuntimeError):
    """Projection phase report creation failed closed without PHI in errors."""


def create_projections_phase_report(
    workbook_path: str | Path,
    *,
    previous_operation: DailyInputOperation,
    operation: DailyInputOperation,
    preflight: DailyInputProjectionPreflight,
    evidence_root: str | Path,
    report_path: str | Path,
    report_id: str,
    source_excel_factory=None,
    input_ledger_excel_factory=None,
    master_excel_factory=None,
    therapist_excel_factory=None,
) -> DailyInputPhaseVerifierReport:
    """Inspect exact saved workbook bytes and exclusively publish a report.

    All expected MASTER and therapist facts must come from the independent
    preflight. The composite diagnostic opens fresh macro-disabled read-only
    Excel instances. Report claims remain untrusted until external attestation.
    """
    if (type(previous_operation) is not DailyInputOperation
            or type(operation) is not DailyInputOperation):
        raise ProjectionsPhaseReportAdapterError("Invalid operation type")
    try:
        obligations = required_evidence_for_transition(previous_operation, operation)
    except PhaseEvidenceContractError as exc:
        raise ProjectionsPhaseReportAdapterError("Invalid phase transition") from exc
    if (operation.phase != DailyInputOperationPhase.PROJECTIONS_VERIFIED
            or obligations != (EvidenceKind.PROJECTIONS_REOPEN,)):
        raise ProjectionsPhaseReportAdapterError("Expected PROJECTIONS_VERIFIED phase transition")
    if type(preflight) is not DailyInputProjectionPreflight:
        raise ProjectionsPhaseReportAdapterError("Independent projection preflight required")
    try:
        day = date.fromisoformat(operation.service_date)
    except ValueError as exc:
        raise ProjectionsPhaseReportAdapterError("Invalid operation service date") from exc
    if (preflight.service_date != day
            or not preflight.validated_changes
            or type(report_id) is not str or not report_id.strip()):
        raise ProjectionsPhaseReportAdapterError("Preflight/operation identity mismatch")

    try:
        root = Path(evidence_root).resolve(strict=True)
        workbook = Path(workbook_path).resolve(strict=True)
        output = Path(report_path).resolve(strict=False)
    except (ValueError, OSError) as exc:
        raise ProjectionsPhaseReportAdapterError("Evidence path resolution failed") from exc
    if not root.is_dir() or not workbook.is_file() or workbook.suffix.casefold() != ".xlsm":
        raise ProjectionsPhaseReportAdapterError("Existing XLSM inside evidence root required")
    try:
        locator = workbook.relative_to(root).as_posix()
        output.relative_to(root)
    except ValueError as exc:
        raise ProjectionsPhaseReportAdapterError("Report and workbook must remain within evidence root") from exc
    if not locator or output == workbook or output.exists():
        raise ProjectionsPhaseReportAdapterError("Evidence artifact/report overlap or output already exists")
    if not output.parent.is_dir():
        raise ProjectionsPhaseReportAdapterError("Report parent directory must already exist")

    initial_hash = file_sha256(workbook)
    if initial_hash != preflight.source_workbook_sha256:
        raise ProjectionsPhaseReportAdapterError("Workbook bytes differ from independent preflight")
    try:
        native = verify_daily_input_projection_preflight_native_readonly(
            workbook,
            preflight=preflight,
            source_excel_factory=source_excel_factory,
            input_ledger_excel_factory=input_ledger_excel_factory,
            master_excel_factory=master_excel_factory,
            therapist_excel_factory=therapist_excel_factory,
        )
    except DailyInputProjectionPreflightNativeError as exc:
        raise ProjectionsPhaseReportAdapterError("Independent native projection readback failed") from exc
    final_hash = file_sha256(workbook)
    if type(native) is not DailyInputProjectionPreflightNativeReadback:
        raise ProjectionsPhaseReportAdapterError("Native projection verifier returned invalid result")
    if (native.source_workbook_sha256 != initial_hash
            or native.native_readback.workbook_sha256 != initial_hash
            or native.source_capture.workbook_sha256 != initial_hash
            or final_hash != initial_hash
            or native.service_date != day):
        raise ProjectionsPhaseReportAdapterError("Workbook bytes/date changed during projection verification")

    projection = native.native_readback
    capture = native.source_capture
    expected_grid = preflight.expected_therapist_grid
    versions = (
        capture.native_excel_version,
        projection.master_excel_version,
        projection.therapist_excel_version,
    )
    if (any(type(value) is not str or not value.strip() for value in versions)
            or len(set(versions)) != 1):
        raise ProjectionsPhaseReportAdapterError("Native Excel versions are missing or inconsistent")
    if (projection.master_projected_rows != len(preflight.expected_projection)
            or projection.therapist_grid_cells_checked != 126
            or projection.therapist_occupied_cells != expected_grid.occupied_cells
            or projection.expected_therapist_assignments != len(expected_grid.expected_assignments)
            or projection.input_ledger_source_changes != len(preflight.validated_changes)
            or projection.input_ledger_active_rows != len(preflight.validated_changes)
            or native.native_source_identity_mirror_rows_checked != preflight.patient_count):
        raise ProjectionsPhaseReportAdapterError("Native projection scope differs from independent preflight")

    contract = VERIFIER_CONTRACTS[EvidenceKind.PROJECTIONS_REOPEN]
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    try:
        report = create_phase_verifier_report(
            report_id=report_id,
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=EvidenceKind.PROJECTIONS_REOPEN.value,
            verifier_id=contract.verifier_id,
            contract_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result="VERIFIED",
            evidence_artifact=locator,
            evidence_sha256=final_hash,
            inspected_scope=list(contract.required_scope),
            native_excel_version=versions[0],
            package_audit_sha256=None,
            prepared_snapshot_sha256=None,
            verified_at_utc=timestamp,
        )
    except PhaseVerifierContractError as exc:
        raise ProjectionsPhaseReportAdapterError("Phase report contract rejected diagnostic") from exc

    try:
        with output.open("x", encoding="utf-8", newline="\\n") as handle:
            handle.write(report.to_json())
            handle.write("\\n")
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError as exc:
        raise ProjectionsPhaseReportAdapterError("Phase report output already exists") from exc
    except OSError as exc:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        raise ProjectionsPhaseReportAdapterError("Phase report publication failed") from exc
    return report
