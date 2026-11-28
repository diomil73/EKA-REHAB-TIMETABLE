"""Non-authorizing native-readback adapter for REPLACEMENTS_REOPEN.

The final workbook's literal DAILY_INPUT and DAILY_REPLACEMENT_FACTS source
cells are reopened through the existing macro-disabled native Excel source
capture. Their independently parsed semantic facts are bound to the prior
pending inventory, operation identity and exact saved workbook digest.
"""
from __future__ import annotations

from datetime import date, datetime, timezone
from pathlib import Path
import os

from .authoritative_commit import file_sha256
from .daily_input_operation import DailyInputOperation, DailyInputOperationPhase
from .daily_input_pending_decision_inventory import (
    PendingDecisionInventoryError, read_pending_decision_inventory_artifact,
)
from .daily_input_phase_evidence import (
    EvidenceKind, PhaseEvidenceContractError, required_evidence_for_transition,
)
from .daily_input_phase_verifier_contract import (
    VERIFIER_CONTRACTS, DailyInputPhaseVerifierReport,
    PhaseVerifierContractError, create_phase_verifier_report,
    decode_phase_verifier_report,
)
from .daily_input_projection_preflight import DailyInputProjectionPreflight
from .daily_input_replacements_reopen import (
    ReplacementsReopenError, build_replacements_reopen_evidence,
    read_replacements_reopen_artifact, write_replacements_reopen_artifact,
)
from .daily_input_source_native_capture import (
    DailyInputSourceNativeCapture, NativeSourceSheetResult, REQUIRED_SOURCE_SHEETS,
    verify_daily_input_source_native_capture_readonly,
)


class ReplacementsReopenPhaseReportAdapterError(RuntimeError):
    """Replacement/suspension phase report creation failed closed."""


def create_replacements_reopen_phase_report(
    workbook_path: str | Path,
    pending_inventory_path: str | Path,
    *,
    previous_operation: DailyInputOperation,
    operation: DailyInputOperation,
    preflight: DailyInputProjectionPreflight,
    evidence_root: str | Path,
    evidence_path: str | Path,
    report_path: str | Path,
    report_id: str,
    source_excel_factory=None,
) -> DailyInputPhaseVerifierReport:
    """Reopen replacement/suspension source facts and publish one report-only report.

    This adapter requires the immediately preceding PENDING_OPERATOR_DECISION
    operation and its exact inventory artifact. It verifies that every pending
    target is now represented by exactly one active replacement or explicit
    treatment suspension. Native Excel reopens the exact final XLSM with macros
    disabled and ReadOnly=True; the existing source capture compares every
    populated DAILY_INPUT and DAILY_REPLACEMENT_FACTS cell with OOXML literals.
    No workbook, VBA, journal or receipt is modified.
    """
    if (type(previous_operation) is not DailyInputOperation
            or type(operation) is not DailyInputOperation
            or type(preflight) is not DailyInputProjectionPreflight):
        raise ReplacementsReopenPhaseReportAdapterError("Exact operation and preflight types required")
    try:
        obligations = required_evidence_for_transition(previous_operation, operation)
    except PhaseEvidenceContractError as exc:
        raise ReplacementsReopenPhaseReportAdapterError("Invalid operation phase transition") from exc
    if (previous_operation.phase != DailyInputOperationPhase.PENDING_OPERATOR_DECISION
            or operation.phase != DailyInputOperationPhase.REPLACEMENTS_RESOLVED
            or EvidenceKind.REPLACEMENTS_REOPEN not in obligations
            or EvidenceKind.OPERATOR_RESOLUTION not in obligations):
        raise ReplacementsReopenPhaseReportAdapterError("Expected pending-decision to replacements-resolved transition")
    try:
        day = date.fromisoformat(operation.service_date)
    except ValueError as exc:
        raise ReplacementsReopenPhaseReportAdapterError("Invalid operation service date") from exc
    if (type(preflight.service_date) is not date or preflight.service_date != day
            or type(report_id) is not str or not report_id.strip()):
        raise ReplacementsReopenPhaseReportAdapterError("Preflight/operation identity mismatch")

    try:
        root = Path(evidence_root).resolve(strict=True)
        workbook = Path(workbook_path).resolve(strict=True)
        inventory_file = Path(pending_inventory_path).resolve(strict=True)
        evidence_output = Path(evidence_path).resolve(strict=False)
        report_output = Path(report_path).resolve(strict=False)
        workbook_locator = workbook.relative_to(root).as_posix()
        inventory_file.relative_to(root)
        evidence_output.relative_to(root)
        report_output.relative_to(root)
    except (OSError, ValueError) as exc:
        raise ReplacementsReopenPhaseReportAdapterError(
            "Workbook and all evidence/report paths must remain within evidence root"
        ) from exc
    if (not root.is_dir() or not workbook.is_file()
            or workbook.suffix.casefold() != ".xlsm"
            or not inventory_file.is_file() or not workbook_locator
            or workbook_locator == "."):
        raise ReplacementsReopenPhaseReportAdapterError("Existing saved XLSM and pending inventory required")
    if (workbook == inventory_file
            or evidence_output in {workbook, inventory_file}
            or report_output in {workbook, inventory_file, evidence_output}
            or evidence_output.exists() or report_output.exists()
            or not evidence_output.parent.is_dir() or not report_output.parent.is_dir()):
        raise ReplacementsReopenPhaseReportAdapterError("Evidence/report overlap or output already exists")

    initial_workbook_hash = file_sha256(workbook)
    try:
        inventory, inventory_artifact = read_pending_decision_inventory_artifact(
            evidence_root=root, artifact_path=inventory_file,
        )
        if (inventory.operation_id != previous_operation.operation_id
                or inventory.operation_source_sha256 != previous_operation.source_sha256
                or inventory.service_date != previous_operation.service_date):
            raise ReplacementsReopenError("Prior inventory does not bind to preceding operation")
        evidence = build_replacements_reopen_evidence(
            previous_operation, operation, preflight, inventory,
        )
    except (PendingDecisionInventoryError, ReplacementsReopenError) as exc:
        raise ReplacementsReopenPhaseReportAdapterError(
            "Pending inventory or final replacement/suspension facts rejected"
        ) from exc
    if (preflight.source_workbook_sha256 != initial_workbook_hash
            or evidence.evidence_workbook_sha256 != initial_workbook_hash):
        raise ReplacementsReopenPhaseReportAdapterError("Final preflight does not bind to exact saved workbook")

    try:
        capture = verify_daily_input_source_native_capture_readonly(
            workbook, excel_factory=source_excel_factory,
        )
    except Exception as exc:
        raise ReplacementsReopenPhaseReportAdapterError(
            "Independent native source-sheet readback failed"
        ) from exc

    final_workbook_hash = file_sha256(workbook)
    if type(capture) is not DailyInputSourceNativeCapture:
        raise ReplacementsReopenPhaseReportAdapterError("Native source capture returned invalid result")
    if (capture.workbook_sha256 != initial_workbook_hash
            or final_workbook_hash != initial_workbook_hash
            or type(capture.native_excel_version) is not str
            or not capture.native_excel_version.strip()
            or type(capture.identity_mirror_rows_checked) is not int
            or capture.identity_mirror_rows_checked != preflight.patient_count
            or type(capture.checked_formula_cells) is not int
            or capture.checked_formula_cells < 0
            or type(capture.sheets) is not tuple
            or any(type(item) is not NativeSourceSheetResult for item in capture.sheets)):
        raise ReplacementsReopenPhaseReportAdapterError("Native source digest/version/count binding mismatch")
    sheet_names = tuple(item.sheet_name for item in capture.sheets)
    if (len(set(sheet_names)) != len(sheet_names)
            or not REQUIRED_SOURCE_SHEETS.issubset(set(sheet_names))
            or ("DAILY_REPLACEMENT_FACTS" not in sheet_names and evidence.replacement_count > 0)):
        raise ReplacementsReopenPhaseReportAdapterError("Native source readback omitted required replacement/suspension scope")

    artifact = None
    report_created = False
    completed = False
    try:
        artifact = write_replacements_reopen_artifact(
            evidence, evidence_root=root, artifact_path=evidence_output,
        )
        reopened, persisted_artifact = read_replacements_reopen_artifact(
            evidence_root=root, artifact_path=evidence_output,
        )
        if reopened != evidence or persisted_artifact != artifact:
            raise ReplacementsReopenPhaseReportAdapterError("Replacement evidence artifact readback mismatch")
        if (file_sha256(workbook) != initial_workbook_hash
                or file_sha256(inventory_file) != inventory_artifact.artifact_sha256):
            raise ReplacementsReopenPhaseReportAdapterError("Workbook or pending inventory changed during native verification")

        contract = VERIFIER_CONTRACTS[EvidenceKind.REPLACEMENTS_REOPEN]
        timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        report = create_phase_verifier_report(
            report_id=report_id,
            operation_id=operation.operation_id,
            source_workbook_sha256=operation.source_sha256,
            service_date=operation.service_date,
            journal_revision=operation.revision,
            journal_phase=operation.phase.value,
            evidence_kind=EvidenceKind.REPLACEMENTS_REOPEN.value,
            verifier_id=contract.verifier_id,
            contract_version=contract.contract_version,
            verification_method=contract.verification_method,
            verification_result="VERIFIED",
            evidence_artifact=artifact.artifact_locator,
            evidence_sha256=artifact.artifact_sha256,
            inspected_scope=list(contract.required_scope),
            native_excel_version=capture.native_excel_version,
            package_audit_sha256=None,
            prepared_snapshot_sha256=None,
            verified_at_utc=timestamp,
        )
        with report_output.open("x", encoding="utf-8", newline="\n") as handle:
            report_created = True
            handle.write(report.to_json())
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        persisted_report = decode_phase_verifier_report(report_output.read_text(encoding="utf-8").strip())
        if persisted_report != report:
            raise ReplacementsReopenPhaseReportAdapterError("Phase report persistence readback mismatch")
        completed = True
    except FileExistsError as exc:
        raise ReplacementsReopenPhaseReportAdapterError("Evidence/report output already exists") from exc
    except (ReplacementsReopenError, PhaseVerifierContractError, OSError, UnicodeError) as exc:
        raise ReplacementsReopenPhaseReportAdapterError("Replacement evidence/report publication failed closed") from exc
    except Exception as exc:
        raise ReplacementsReopenPhaseReportAdapterError("Replacement phase report failed closed") from exc
    finally:
        if not completed:
            if report_created:
                try:
                    report_output.unlink(missing_ok=True)
                except OSError:
                    pass
            if artifact is not None:
                try:
                    evidence_output.unlink(missing_ok=True)
                except OSError:
                    pass
    return report
