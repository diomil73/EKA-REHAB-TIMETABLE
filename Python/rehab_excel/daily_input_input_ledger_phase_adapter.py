"""Non-authorizing INPUT_AND_LEDGER native diagnostic phase-report adapter.

Produces a canonical report only after exact OOXML + native Excel COM readback
of a saved disposable XLSM under an evidence root. This report is NOT a
signature, trusted attestation, receipt, journal append or permission for Apply.
"""
from __future__ import annotations

from datetime import date, datetime, timezone
from pathlib import Path
import os

from .authoritative_commit import file_sha256
from .daily_input_input_ledger_native import (
    InputLedgerNativeReadbackError,
    verify_input_ledger_native_readonly,
)
from .daily_input_operation import DailyInputOperation, DailyInputOperationPhase
from .daily_input_phase_evidence import (
    EvidenceKind, PhaseEvidenceContractError, required_evidence_for_transition,
)
from .daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS, DailyInputPhaseVerifierReport,
    PhaseVerifierContractError, create_phase_verifier_report,
)


class InputLedgerPhaseReportAdapterError(RuntimeError):
    """Report creation rejected, without exposing patient-identifying data."""


def create_input_ledger_phase_report(
    workbook_path: str | Path,
    *,
    previous_operation: DailyInputOperation,
    operation: DailyInputOperation,
    validated_changes: tuple,
    patients,
    sessions,
    evidence_root: str | Path,
    report_path: str | Path,
    report_id: str,
    therapist_names=(),
    excel_factory=None,
) -> DailyInputPhaseVerifierReport:
    """Inspect saved source bytes, and exclusively publish a readback report.

    The workbook is the exact evidence artifact; report provenance is untrusted.
    The expected changes MUST come from an earlier independent preflight.
    No positive transition/Apply authorization is possible through this adapter.
    """
    if (type(previous_operation) is not DailyInputOperation
            or type(operation) is not DailyInputOperation):
        raise InputLedgerPhaseReportAdapterError("Invalid operation type")
    try:
        obligations = required_evidence_for_transition(previous_operation, operation)
    except PhaseEvidenceContractError as exc:
        raise InputLedgerPhaseReportAdapterError("Invalid phase transition") from exc
    if (operation.phase != DailyInputOperationPhase.INPUT_AND_LEDGER_COMMITTED
            or obligations != (EvidenceKind.INPUT_AND_LEDGER_REOPEN,)):
        raise InputLedgerPhaseReportAdapterError("Expected INPUT_AND_LEDGER phase transition")

    try:
        day = date.fromisoformat(operation.service_date)
    except ValueError as exc:
        raise InputLedgerPhaseReportAdapterError("Invalid operation service date") from exc
    if type(report_id) is not str or not report_id.strip():
        raise InputLedgerPhaseReportAdapterError("Invalid report ID")
    if type(validated_changes) is not tuple or not validated_changes:
        raise InputLedgerPhaseReportAdapterError("Non-empty validated preflight facts required")

    try:
        root = Path(evidence_root).resolve(strict=True)
        workbook = Path(workbook_path).resolve(strict=True)
        output = Path(report_path).resolve(strict=False)
    except (ValueError, OSError) as exc:
        raise InputLedgerPhaseReportAdapterError("Evidence path resolution failed") from exc
    if not root.is_dir() or not workbook.is_file() or workbook.suffix.casefold() != ".xlsm":
        raise InputLedgerPhaseReportAdapterError("Existing XLSM inside evidence root required")
    try:
        locator = workbook.relative_to(root).as_posix()
        output.relative_to(root)
    except ValueError as exc:
        raise InputLedgerPhaseReportAdapterError("Report and workbook must remain within evidence root") from exc
    if not locator or output == workbook or output.exists():
        raise InputLedgerPhaseReportAdapterError("Evidence artifact/report overlap or output already exists")
    if not output.parent.is_dir():
        raise InputLedgerPhaseReportAdapterError("Report parent directory must already exist")

    initial_hash = file_sha256(workbook)
    try:
        native = verify_input_ledger_native_readonly(
            workbook, service_date=day, validated_changes=validated_changes,
            patients=patients, sessions=sessions, therapist_names=therapist_names,
            excel_factory=excel_factory,
        )
    except InputLedgerNativeReadbackError as exc:
        raise InputLedgerPhaseReportAdapterError("Independent input/ledger readback failed") from exc
    final_hash = file_sha256(workbook)
    if native.workbook_sha256 != initial_hash or final_hash != initial_hash:
        raise InputLedgerPhaseReportAdapterError("Workbook bytes changed during phase-report verification")
    if native.service_date != day or native.source_change_count != len(validated_changes):
        raise InputLedgerPhaseReportAdapterError("Native verification result inconsistent with preflight")
    if native.active_ledger_count != len(validated_changes):
        raise InputLedgerPhaseReportAdapterError("Native ledger result inconsistent with preflight")
    if type(native.native_excel_version) is not str or not native.native_excel_version.strip():
        raise InputLedgerPhaseReportAdapterError("Native Excel version missing")

    contract = VERIFIER_CONTRACTS[EvidenceKind.INPUT_AND_LEDGER_REOPEN]
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    try:
        report = create_phase_verifier_report(
            report_id=report_id,
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=EvidenceKind.INPUT_AND_LEDGER_REOPEN.value,
            verifier_id=contract.verifier_id,
            contract_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result="VERIFIED",
            evidence_artifact=locator,
            evidence_sha256=final_hash,
            inspected_scope=list(contract.required_scope),
            native_excel_version=native.native_excel_version.strip(),
            package_audit_sha256=None,
            prepared_snapshot_sha256=None,
            verified_at_utc=timestamp,
        )
    except PhaseVerifierContractError as exc:
        raise InputLedgerPhaseReportAdapterError("Phase report contract rejected diagnostic") from exc

    try:
        with output.open("x", encoding="utf-8", newline="\n") as handle:
            handle.write(report.to_json())
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError as exc:
        raise InputLedgerPhaseReportAdapterError("Phase report output already exists") from exc
    except OSError as exc:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        raise InputLedgerPhaseReportAdapterError("Phase report publication failed") from exc
    return report
