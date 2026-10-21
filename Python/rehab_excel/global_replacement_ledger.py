from __future__ import annotations

from typing import Iterable

from rehab_core.global_replacement_end_to_end import GlobalReplacementAssignmentResult
from rehab_core.models import Patient, Session

from .daily_change_ledger_vba import DailyChangeLedgerRecord
from .patient_centric_preview import PreviewBinding


class GlobalReplacementLedgerError(RuntimeError):
    """Raised when a global replacement plan cannot be bound to the daily ledger."""


def build_global_replacement_ledger_records(
    result: GlobalReplacementAssignmentResult,
    *,
    sessions: Iterable[Session],
    patients: Iterable[Patient],
    bindings: Iterable[PreviewBinding],
) -> tuple[DailyChangeLedgerRecord, ...]:
    """Bind a finalized global plan to exact THERAPIST_DAILY source/destination cells.

    This function is mutation-free. It requires one preview binding per assignment
    and refuses unresolved global plans, because a partially resolved day must never
    be presented as a final reversible ledger transaction.
    """

    if result.unresolved_request_ids:
        raise GlobalReplacementLedgerError(
            "Cannot build a final daily ledger while global replacements remain unresolved"
        )

    session_by_id = {session.session_id: session for session in sessions}
    patient_by_id = {patient.patient_id: patient for patient in patients}
    binding_by_session: dict[str, PreviewBinding] = {}
    for binding in bindings:
        if binding.session_id in binding_by_session:
            raise GlobalReplacementLedgerError(
                f"Duplicate preview binding for session {binding.session_id!r}"
            )
        binding_by_session[binding.session_id] = binding

    records: list[DailyChangeLedgerRecord] = []
    for assignment in result.assignments:
        target = session_by_id.get(assignment.target_session_id)
        if target is None:
            raise GlobalReplacementLedgerError(
                f"Missing target session {assignment.target_session_id!r}"
            )
        patient = patient_by_id.get(assignment.patient_id)
        if patient is None:
            raise GlobalReplacementLedgerError(
                f"Missing patient {assignment.patient_id!r}"
            )
        binding = binding_by_session.get(assignment.target_session_id)
        if binding is None:
            raise GlobalReplacementLedgerError(
                f"Missing preview binding for session {assignment.target_session_id!r}"
            )
        if binding.patient_id != assignment.patient_id:
            raise GlobalReplacementLedgerError(
                f"Preview binding patient mismatch for session {assignment.target_session_id!r}"
            )
        if not binding.effective_cell:
            raise GlobalReplacementLedgerError(
                f"Replacement session {assignment.target_session_id!r} has no effective cell"
            )

        records.append(
            DailyChangeLedgerRecord(
                change_id=assignment.replacement_id,
                change_date=assignment.replacement_date,
                patient_id=assignment.patient_id,
                patient_name=patient.display_name,
                original_provider=assignment.original_therapist_id,
                original_time=target.start_time,
                effective_provider=assignment.replacement_therapist_id,
                effective_time=assignment.replacement_time,
                original_cell=binding.original_cell,
                effective_cell=binding.effective_cell,
                reason=assignment.reason or "προσωρινή ανακατανομή",
            )
        )

    return tuple(sorted(records, key=lambda item: item.change_id))
