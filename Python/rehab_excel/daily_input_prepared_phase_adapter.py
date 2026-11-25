"""Adapter from the existing PREPARED native readback to a phase report.

The report records the result of the existing independent diagnostic. It is
not a trusted attestation and cannot satisfy phase authorization on its own.
"""
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
import os

from .authoritative_commit import file_sha256
from .daily_input_phase_evidence import EvidenceKind
from .daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS,
    DailyInputPhaseVerifierReport,
    PhaseVerifierContractError,
    create_phase_verifier_report,
)
from .daily_input_operation import DailyInputOperation


class PreparedPhaseReportAdapterError(RuntimeError):
    """PREPARED readback or phase-report publication failed closed."""


def create_prepared_snapshot_phase_report(
    workbook_path: str | Path,
    handoff: dict[str, object],
    payload_path: str | Path,
    *,
    evidence_root: str | Path,
    report_path: str | Path,
    report_id: str,
    excel_factory=None,
) -> DailyInputPhaseVerifierReport:
    """Run the existing native PREPARED verifier and publish its bound report.

    The workbook itself is the evidence artifact. It must be inside evidence_root
    and its exact saved bytes are hashed after native read-only reopen. The report
    is created exclusively and never replaces existing evidence. This adapter
    does not sign/attest the report and does not authorize a journal transition.
    """
    from tools.verify_daily_input_journal_prepared import (
        PreparedJournalReadbackError,
        verify_prepared_reopen,
    )

    expected_handoff_fields = {
        "status", "operation_id", "service_date", "source_sha256",
        "payload_sha256", "existing_journal_rows", "journal_prefix_sha256",
        "prepared_row",
    }
    if type(handoff) is not dict or set(handoff) != expected_handoff_fields:
        raise PreparedPhaseReportAdapterError("PREPARED handoff schema invalid")
    workbook = Path(workbook_path).resolve(strict=True)
    payload = Path(payload_path).resolve(strict=True)
    root = Path(evidence_root).resolve(strict=True)
    output = Path(report_path).resolve(strict=False)
    if not workbook.is_file() or workbook.suffix.casefold() != ".xlsm":
        raise PreparedPhaseReportAdapterError("PREPARED evidence must be an existing XLSM file")
    if not payload.is_file() or payload == workbook:
        raise PreparedPhaseReportAdapterError("PREPARED payload is unavailable or overlaps workbook")
    if not root.is_dir():
        raise PreparedPhaseReportAdapterError("PREPARED evidence root is not a directory")
    try:
        relative_artifact = workbook.relative_to(root).as_posix()
        output.relative_to(root)
    except ValueError:
        raise PreparedPhaseReportAdapterError("Workbook and report must remain inside evidence root") from None
    if output in {workbook, payload} or output.exists():
        raise PreparedPhaseReportAdapterError("PREPARED phase report output already exists or overlaps input")
    if not report_id or type(report_id) is not str:
        raise PreparedPhaseReportAdapterError("PREPARED report ID is invalid")

    try:
        operation = DailyInputOperation(
            operation_id=handoff["operation_id"],
            source_sha256=handoff["source_sha256"],
            service_date=handoff["service_date"],
        )
    except (TypeError, ValueError, AttributeError):
        raise PreparedPhaseReportAdapterError("PREPARED operation identity is invalid") from None
    contract = VERIFIER_CONTRACTS[EvidenceKind.PREPARED_SNAPSHOT]
    before_snapshot_digest = file_sha256(workbook)
    try:
        native_version = verify_prepared_reopen(
            workbook, handoff, payload, native_excel=True, excel_factory=excel_factory,
        )
    except PreparedJournalReadbackError:
        raise PreparedPhaseReportAdapterError("Independent PREPARED native readback failed") from None
    if type(native_version) is not str or not native_version.strip():
        raise PreparedPhaseReportAdapterError("Native Excel version was not reported")

    snapshot_digest = file_sha256(workbook)
    if snapshot_digest != before_snapshot_digest:
        raise PreparedPhaseReportAdapterError("PREPARED workbook bytes changed during verifier adapter run")
    if file_sha256(payload) != handoff["payload_sha256"]:
        raise PreparedPhaseReportAdapterError("PREPARED payload changed after native readback")
    artifact = root.joinpath(*PurePosixPath(relative_artifact).parts).resolve(strict=True)
    try:
        artifact.relative_to(root)
    except ValueError:
        raise PreparedPhaseReportAdapterError("PREPARED evidence artifact escapes evidence root") from None
    if artifact != workbook or file_sha256(artifact) != snapshot_digest:
        raise PreparedPhaseReportAdapterError("PREPARED evidence artifact does not match reopened workbook bytes")

    verified_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    try:
        report = create_phase_verifier_report(
            report_id=report_id,
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=EvidenceKind.PREPARED_SNAPSHOT.value,
            verifier_id=contract.verifier_id,
            contract_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result="VERIFIED",
            evidence_artifact=relative_artifact,
            evidence_sha256=snapshot_digest,
            inspected_scope=list(contract.required_scope),
            native_excel_version=native_version.strip(),
            package_audit_sha256=None,
            prepared_snapshot_sha256=snapshot_digest,
            verified_at_utc=verified_at,
        )
    except PhaseVerifierContractError as exc:
        raise PreparedPhaseReportAdapterError("PREPARED phase report contract rejected readback") from exc

    try:
        output.parent.mkdir(parents=False, exist_ok=True)
        with output.open("x", encoding="utf-8", newline="\n") as handle:
            handle.write(report.to_json())
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
    except FileExistsError:
        raise PreparedPhaseReportAdapterError("PREPARED phase report output already exists") from None
    except OSError:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        raise PreparedPhaseReportAdapterError("PREPARED phase report publication failed") from None
    return report
