from datetime import date, time

import pytest

from rehab_core.daily_state import DailySessionStatus
from rehab_core.global_replacement_end_to_end import GlobalReplacementAssignmentResult
from rehab_core.models import Patient, ReplacementAssignment, Session
from rehab_excel.global_replacement_ledger import (
    GlobalReplacementLedgerError,
    build_global_replacement_ledger_records,
)
from rehab_excel.patient_centric_preview import PreviewBinding


TODAY = date(2026, 10, 21)


def _session() -> Session:
    return Session(
        session_id="S1",
        patient_id="P1",
        therapist_id="T-OLD",
        session_date=TODAY,
        start_time=time(9, 0),
    )


def _patient() -> Patient:
    return Patient(patient_id="P1", display_name="Ασθενής Α")


def _assignment() -> ReplacementAssignment:
    return ReplacementAssignment(
        replacement_id="R1",
        target_session_id="S1",
        patient_id="P1",
        original_therapist_id="T-OLD",
        replacement_therapist_id="T-NEW",
        replacement_date=TODAY,
        replacement_time=time(10, 0),
        reason="global replacement plan",
    )


def _binding(effective_cell: str | None = "E8") -> PreviewBinding:
    return PreviewBinding(
        session_id="S1",
        patient_id="P1",
        status=DailySessionStatus.REPLACED,
        original_cell="C7",
        effective_cell=effective_cell,
    )


def test_builds_reversible_ledger_record_from_global_assignment_and_preview_binding():
    result = GlobalReplacementAssignmentResult(assignments=(_assignment(),), unresolved_request_ids=())

    records = build_global_replacement_ledger_records(
        result,
        sessions=(_session(),),
        patients=(_patient(),),
        bindings=(_binding(),),
    )

    assert len(records) == 1
    record = records[0]
    assert record.change_id == "R1"
    assert record.patient_name == "Ασθενής Α"
    assert record.original_provider == "T-OLD"
    assert record.original_time == time(9, 0)
    assert record.effective_provider == "T-NEW"
    assert record.effective_time == time(10, 0)
    assert record.original_cell == "C7"
    assert record.effective_cell == "E8"


def test_rejects_unresolved_global_plan_before_ledger_creation():
    result = GlobalReplacementAssignmentResult(assignments=(_assignment(),), unresolved_request_ids=("S2",))

    with pytest.raises(GlobalReplacementLedgerError, match="remain unresolved"):
        build_global_replacement_ledger_records(
            result,
            sessions=(_session(),),
            patients=(_patient(),),
            bindings=(_binding(),),
        )


def test_rejects_missing_effective_destination_cell():
    result = GlobalReplacementAssignmentResult(assignments=(_assignment(),), unresolved_request_ids=())

    with pytest.raises(GlobalReplacementLedgerError, match="has no effective cell"):
        build_global_replacement_ledger_records(
            result,
            sessions=(_session(),),
            patients=(_patient(),),
            bindings=(_binding(None),),
        )


def test_rejects_binding_patient_mismatch():
    result = GlobalReplacementAssignmentResult(assignments=(_assignment(),), unresolved_request_ids=())
    bad_binding = PreviewBinding(
        session_id="S1",
        patient_id="P-OTHER",
        status=DailySessionStatus.REPLACED,
        original_cell="C7",
        effective_cell="E8",
    )

    with pytest.raises(GlobalReplacementLedgerError, match="patient mismatch"):
        build_global_replacement_ledger_records(
            result,
            sessions=(_session(),),
            patients=(_patient(),),
            bindings=(bad_binding,),
        )
