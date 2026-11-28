"""Non-authorizing PENDING_DECISION_INVENTORY native-readback phase adapter.

The inventory is independently reconstructed from the preflight, reopened
from exact canonical artifact bytes, and bound to a fresh macro-disabled
native read-only Excel verification of the exact saved XLSM. The report is
not trusted attestation, a receipt, a journal append, or Apply authorization.
"""
from __future__ import annotations

from datetime import date, datetime, timezone
from pathlib import Path
import os

from .authoritative_commit import file_sha256
from .daily_input_operation import DailyInputOperation, DailyInputOperationPhase
from .daily_input_pending_decision_inventory import (
    PendingDecisionInventory,
    PendingDecisionInventoryError,
    build_pending_decision_inventory,
    read_pending_decision_inventory_artifact,
)
from .daily_input_phase_evidence import (
    EvidenceKind, PhaseEvidenceContractError, required_evidence_for_transition,
)
from .daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS,
    DailyInputPhaseVerifierReport,
    PhaseVerifierContractError,
    create_phase_verifier_report,
)
from .daily_input_projection_native import DailyInputProjectionNativeReadback
from .daily_input_projection_preflight import DailyInputProjectionPreflight
from .daily_input_projection_preflight_native import (
    DailyInputProjectionPreflightNativeError,
    DailyInputProjectionPreflightNativeReadback,
    verify_daily_input_projection_preflight_native_readonly,
)
from .daily_input_source_native_capture import (
    DailyInputSourceNativeCapture, REQUIRED_SOURCE_SHEETS,
)


class PendingDecisionPhaseReportAdapterError(RuntimeError):
    """Pending decision report creation failed closed without exposing PHI."""


def create_pending_decision_phase_report(
    workbook_path: str | Path,
    inventory_path: str | Path,
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
    """Reopen and bind the exact pending inventory and saved XLSM; publish once.

    The expected inventory is rebuilt from the same independent preflight and
    must exactly equal the reopened canonical artifact. Native Excel then
    independently reopens the exact saved XLSM (macros disabled/read-only)
    through the existing composite source/input/ledger/projection verifier.
    No workbook or journal is modified.
    """
    if (type(previous_operation) is not DailyInputOperation
            or type(operation) is not DailyInputOperation
            or type(preflight) is not DailyInputProjectionPreflight):
        raise PendingDecisionPhaseReportAdapterError("Exact operation and preflight types required")
    try:
        obligations = required_evidence_for_transition(previous_operation, operation)
    except PhaseEvidenceContractError as exc:
        raise PendingDecisionPhaseReportAdapterError("Invalid operation phase transition") from exc
    if (operation.phase != DailyInputOperationPhase.PENDING_OPERATOR_DECISION
            or obligations != (EvidenceKind.PENDING_DECISION_INVENTORY,)):
        raise PendingDecisionPhaseReportAdapterError("Expected PENDING_OPERATOR_DECISION phase transition")
    try:
        day = date.fromisoformat(operation.service_date)
    except ValueError as exc:
        raise PendingDecisionPhaseReportAdapterError("Invalid operation service date") from exc
    if (type(preflight.service_date) is not date
            or preflight.service_date != day
            or type(report_id) is not str or not report_id.strip()):
        raise PendingDecisionPhaseReportAdapterError("Preflight/operation identity mismatch")

    try:
        root = Path(evidence_root).resolve(strict=True)
        workbook = Path(workbook_path).resolve(strict=True)
        inventory_file = Path(inventory_path).resolve(strict=True)
        output = Path(report_path).resolve(strict=False)
        workbook_locator = workbook.relative_to(root).as_posix()
        inventory_locator = inventory_file.relative_to(root).as_posix()
        output.relative_to(root)
    except (OSError, ValueError) as exc:
        raise PendingDecisionPhaseReportAdapterError(
            "Workbook, inventory and report must remain within evidence root"
        ) from exc
    if (not root.is_dir() or not workbook.is_file()
            or workbook.suffix.casefold() != ".xlsm"
            or not inventory_file.is_file()
            or not inventory_locator or inventory_locator == "."
            or not workbook_locator or workbook_locator == "."):
        raise PendingDecisionPhaseReportAdapterError("Existing XLSM and inventory artifact required")
    if (workbook == inventory_file or output in {workbook, inventory_file}
            or output.exists() or not output.parent.is_dir()):
        raise PendingDecisionPhaseReportAdapterError("Evidence/report overlap or report output already exists")

    initial_workbook_hash = file_sha256(workbook)
    try:
        inventory, artifact = read_pending_decision_inventory_artifact(
            evidence_root=root, artifact_path=inventory_file,
        )
        expected_inventory = build_pending_decision_inventory(
            previous_operation, operation, preflight,
        )
    except PendingDecisionInventoryError as exc:
        raise PendingDecisionPhaseReportAdapterError("Pending inventory artifact/preflight rejected") from exc
    if inventory != expected_inventory:
        raise PendingDecisionPhaseReportAdapterError("Reopened pending inventory differs from independent preflight")
    if (inventory.operation_id != operation.operation_id
            or inventory.operation_source_sha256 != operation.source_sha256
            or inventory.service_date != operation.service_date
            or inventory.evidence_workbook_sha256 != initial_workbook_hash
            or preflight.source_workbook_sha256 != initial_workbook_hash):
        raise PendingDecisionPhaseReportAdapterError("Pending inventory/workbook operation binding mismatch")

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
        raise PendingDecisionPhaseReportAdapterError("Independent native workbook readback failed") from exc
    except Exception as exc:
        raise PendingDecisionPhaseReportAdapterError("Native pending inventory binding failed closed") from exc

    final_workbook_hash = file_sha256(workbook)
    if type(native) is not DailyInputProjectionPreflightNativeReadback:
        raise PendingDecisionPhaseReportAdapterError("Native verifier returned invalid result")
    if (type(native.preflight) is not DailyInputProjectionPreflight
            or native.preflight is not preflight
            or type(native.native_readback) is not DailyInputProjectionNativeReadback
            or type(native.source_capture) is not DailyInputSourceNativeCapture):
        raise PendingDecisionPhaseReportAdapterError("Native verifier evidence components are invalid")

    projection = native.native_readback
    capture = native.source_capture
    versions = (
        capture.native_excel_version,
        projection.master_excel_version,
        projection.therapist_excel_version,
    )
    sheet_names = tuple(item.sheet_name for item in capture.sheets)
    expected_grid = preflight.expected_therapist_grid
    counts = (
        projection.master_projected_rows,
        projection.therapist_grid_cells_checked,
        projection.therapist_occupied_cells,
        projection.expected_therapist_assignments,
        projection.input_ledger_source_changes,
        projection.input_ledger_active_rows,
        native.native_source_identity_mirror_rows_checked,
    )
    if (native.source_workbook_sha256 != initial_workbook_hash
            or projection.workbook_sha256 != initial_workbook_hash
            or capture.workbook_sha256 != initial_workbook_hash
            or final_workbook_hash != initial_workbook_hash
            or native.service_date != day
            or projection.service_date != day
            or any(type(value) is not str or not value.strip() for value in versions)
            or len(set(versions)) != 1
            or len(set(sheet_names)) != len(sheet_names)
            or not REQUIRED_SOURCE_SHEETS.issubset(set(sheet_names))
            or any(type(value) is not int or value < 0 for value in counts)
            or projection.master_projected_rows != len(preflight.expected_projection)
            or projection.therapist_grid_cells_checked != 126
            or projection.therapist_occupied_cells != expected_grid.occupied_cells
            or projection.expected_therapist_assignments != len(expected_grid.expected_assignments)
            or projection.input_ledger_source_changes != len(preflight.validated_changes)
            or projection.input_ledger_active_rows != len(preflight.validated_changes)
            or native.native_source_identity_mirror_rows_checked != preflight.patient_count):
        raise PendingDecisionPhaseReportAdapterError("Native readback digest/date/scope/version mismatch")
    if (file_sha256(inventory_file) != artifact.artifact_sha256
            or file_sha256(workbook) != initial_workbook_hash):
        raise PendingDecisionPhaseReportAdapterError("Evidence bytes changed during native verification")

    contract = VERIFIER_CONTRACTS[EvidenceKind.PENDING_DECISION_INVENTORY]
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    try:
        report = create_phase_verifier_report(
            report_id=report_id,
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=EvidenceKind.PENDING_DECISION_INVENTORY.value,
            verifier_id=contract.verifier_id,
            contract_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result="VERIFIED",
            evidence_artifact=artifact.artifact_locator,
            evidence_sha256=artifact.artifact_sha256,
            inspected_scope=list(contract.required_scope),
            native_excel_version=versions[0],
            package_audit_sha256=None,
            prepared_snapshot_sha256=None,
            verified_at_utc=timestamp,
        )
    except PhaseVerifierContractError as exc:
        raise PendingDecisionPhaseReportAdapterError("Phase report contract rejected native evidence") from exc

    created = False
    try:
        with output.open("x", encoding="utf-8", newline="\n") as handle:
            created = True
            handle.write(report.to_json())
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError as exc:
        raise PendingDecisionPhaseReportAdapterError("Phase report output already exists") from exc
    except OSError as exc:
        if created:
            try:
                output.unlink(missing_ok=True)
            except OSError:
                pass
        raise PendingDecisionPhaseReportAdapterError("Phase report publication failed") from exc

    try:
        persisted = output.read_text(encoding="utf-8")
        from .daily_input_phase_verifier_contract import decode_phase_verifier_report
        if decode_phase_verifier_report(persisted.strip()) != report:
            raise PendingDecisionPhaseReportAdapterError("Phase report readback mismatch")
    except (OSError, UnicodeError, PhaseVerifierContractError) as exc:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        raise PendingDecisionPhaseReportAdapterError("Phase report persistence readback failed") from exc
    return report
