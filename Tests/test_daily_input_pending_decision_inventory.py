from dataclasses import replace
from datetime import date, time
from hashlib import sha256
import json

import pytest

from rehab_core.daily_state import DailySessionState, DailySessionStatus
from rehab_core.models import Session, ReplacementProviderKind
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_pending_decision_inventory import (
    PendingDecisionInventoryError,
    build_pending_decision_inventory,
    validate_pending_decision_inventory,
    write_pending_decision_inventory,
)
from rehab_excel.daily_input_projection_preflight import DailyInputProjectionPreflight


DAY = date(2026, 11, 27)


def fixture():
    initial = DailyInputOperation("pending-demo-001", "a" * 64, DAY.isoformat())
    projected = advance_daily_input_operation(initial, Phase.INPUT_AND_LEDGER_COMMITTED)
    projected = advance_daily_input_operation(projected, Phase.PROJECTIONS_VERIFIED)
    pending = advance_daily_input_operation(projected, Phase.PENDING_OPERATOR_DECISION)

    sessions = (
        Session("s-active", "p-1", "t-1", DAY, time(8, 30), "ΦΘ"),
        Session("s-patient-absent", "p-2", "t-2", DAY, time(9, 15), "ΦΘ"),
        Session("s-therapist-absent", "p-3", "t-3", DAY, time(10, 0), "ΦΘ", True),
        Session("s-replaced", "p-4", "t-4", DAY, time(10, 45), "ΦΘ"),
        Session("s-cancelled", "p-5", "t-5", DAY, time(11, 30), "ΦΘ"),
    )
    states = (
        DailySessionState(
            "s-active", "p-1", DAY, DailySessionStatus.ACTIVE,
            "t-1", time(8, 30), "t-1", time(8, 30),
            effective_provider_kind=ReplacementProviderKind.THERAPIST,
        ),
        DailySessionState(
            "s-patient-absent", "p-2", DAY, DailySessionStatus.PATIENT_ABSENT,
            "t-2", time(9, 15), None, None,
        ),
        DailySessionState(
            "s-therapist-absent", "p-3", DAY, DailySessionStatus.THERAPIST_ABSENT,
            "t-3", time(10, 0), None, None, reason="synthetic absence",
        ),
        DailySessionState(
            "s-replaced", "p-4", DAY, DailySessionStatus.REPLACED,
            "t-4", time(10, 45), "t-9", time(10, 45),
            replacement_id="replacement-1",
            effective_provider_kind=ReplacementProviderKind.THERAPIST,
        ),
        DailySessionState(
            "s-cancelled", "p-5", DAY, DailySessionStatus.CANCELLED,
            "t-5", time(11, 30), None, None,
            cancellation_id="cancel-1",
        ),
    )
    preflight = DailyInputProjectionPreflight(
        service_date=DAY, baseline_date=date(2026, 1, 5),
        source_workbook_sha256="c" * 64, patients=(), settings=None,
        stable_base_entries=(), effective_base_entries=(), recurring_changes=(),
        patient_exits=(), sessions=sessions, daily_input=None, replacements=(),
        states=states, students=(), expected_projection=(),
        expected_therapist_grid=None, approved_shared_memberships=(),
    )
    return projected, pending, preflight


def test_inventory_contains_every_and_only_unresolved_therapist_absence():
    previous, operation, preflight = fixture()
    inventory = build_pending_decision_inventory(previous, operation, preflight)

    assert inventory.operation_id == operation.operation_id
    assert inventory.operation_source_sha256 == operation.source_sha256
    assert inventory.evidence_workbook_sha256 == preflight.source_workbook_sha256
    assert inventory.operation_source_sha256 != inventory.evidence_workbook_sha256
    assert inventory.total_session_count == 5
    assert inventory.active_count == 1
    assert inventory.patient_absent_count == 1
    assert inventory.therapist_absent_count == 1
    assert inventory.replaced_count == 1
    assert inventory.cancelled_count == 1
    assert len(inventory.targets) == 1
    target = inventory.targets[0]
    assert (
        target.session_id, target.patient_id, target.original_therapist_id,
        target.original_time, target.treatment, target.robotic,
    ) == ("s-therapist-absent", "p-3", "t-3", "10:00", "ΦΘ", True)
    assert "patient_name" not in inventory.to_json()
    assert "provider_options" not in inventory.to_json()


def test_inventory_is_deterministic_and_digest_covers_full_payload():
    previous, operation, preflight = fixture()
    first = build_pending_decision_inventory(previous, operation, preflight)
    reordered = replace(
        preflight,
        sessions=tuple(reversed(preflight.sessions)),
        states=tuple(reversed(preflight.states)),
    )
    second = build_pending_decision_inventory(previous, operation, reordered)
    assert first.to_json() == second.to_json()

    payload = json.loads(first.to_json())
    digest = payload.pop("inventory_sha256")
    canonical = json.dumps(
        payload, ensure_ascii=True, allow_nan=False,
        sort_keys=True, separators=(",", ":"),
    ).encode("ascii")
    assert digest == sha256(canonical).hexdigest()


def test_inventory_rejects_non_pending_transition_and_wrong_operation_date():
    previous, operation, preflight = fixture()
    wrong_phase = advance_daily_input_operation(previous, Phase.REPLACEMENTS_RESOLVED)
    with pytest.raises(PendingDecisionInventoryError, match="PENDING_OPERATOR_DECISION"):
        build_pending_decision_inventory(previous, wrong_phase, preflight)
    wrong_date = replace(operation, service_date="2026-11-28")
    with pytest.raises(PendingDecisionInventoryError, match="service date"):
        build_pending_decision_inventory(previous, wrong_date, preflight)


@pytest.mark.parametrize("mutation, message", [
    ("missing-state", "exact dated session set"),
    ("duplicate-state", "Duplicate preflight daily-state"),
    ("patient-drift", "identity differs"),
    ("time-drift", "identity differs"),
    ("wrong-date", "wrong-date preflight session"),
])
def test_inventory_fails_closed_on_state_or_session_drift(mutation, message):
    previous, operation, preflight = fixture()
    if mutation == "missing-state":
        changed = replace(preflight, states=preflight.states[:-1])
    elif mutation == "duplicate-state":
        changed = replace(preflight, states=preflight.states + (preflight.states[0],))
    elif mutation == "patient-drift":
        changed_state = replace(preflight.states[2], patient_id="different-patient")
        changed = replace(preflight, states=preflight.states[:2] + (changed_state,) + preflight.states[3:])
    elif mutation == "time-drift":
        changed_state = replace(preflight.states[2], original_time=time(10, 1))
        changed = replace(preflight, states=preflight.states[:2] + (changed_state,) + preflight.states[3:])
    else:
        changed_session = replace(preflight.sessions[0], session_date=date(2026, 11, 28))
        changed = replace(preflight, sessions=(changed_session,) + preflight.sessions[1:])
    with pytest.raises(PendingDecisionInventoryError, match=message):
        build_pending_decision_inventory(previous, operation, changed)


def test_pending_phase_cannot_have_empty_inventory():
    previous, operation, preflight = fixture()
    active_only = replace(
        preflight,
        states=tuple(
            replace(state, status=DailySessionStatus.ACTIVE,
                    effective_therapist_id=state.original_therapist_id,
                    effective_time=state.original_time,
                    replacement_id=None)
            for state in preflight.states
        ),
    )
    with pytest.raises(PendingDecisionInventoryError, match="at least one unresolved target"):
        build_pending_decision_inventory(previous, operation, active_only)


def test_inventory_artifact_is_exclusive_and_byte_bound(tmp_path):
    previous, operation, preflight = fixture()
    inventory = build_pending_decision_inventory(previous, operation, preflight)
    root = tmp_path / "evidence"
    root.mkdir()
    nested = root / "operation"
    nested.mkdir()
    output = nested / "pending-inventory.json"

    artifact = write_pending_decision_inventory(
        inventory, evidence_root=root, artifact_path=output,
    )
    payload = output.read_bytes()
    assert artifact.artifact_locator == "operation/pending-inventory.json"
    assert artifact.artifact_sha256 == sha256(payload).hexdigest()
    assert artifact.byte_count == len(payload)
    assert json.loads(payload)["inventory_sha256"] == inventory.inventory_sha256

    with pytest.raises(PendingDecisionInventoryError, match="already exists"):
        write_pending_decision_inventory(
            inventory, evidence_root=root, artifact_path=output,
        )
    assert output.read_bytes() == payload


def test_inventory_artifact_rejects_path_outside_evidence_root(tmp_path):
    previous, operation, preflight = fixture()
    inventory = build_pending_decision_inventory(previous, operation, preflight)
    root = tmp_path / "evidence"
    root.mkdir()
    outside = tmp_path / "outside.json"
    with pytest.raises(PendingDecisionInventoryError, match="escapes evidence root"):
        write_pending_decision_inventory(
            inventory, evidence_root=root, artifact_path=outside,
        )
    assert not outside.exists()


def test_artifact_writer_revalidates_caller_constructed_inventory():
    previous, operation, preflight = fixture()
    inventory = build_pending_decision_inventory(previous, operation, preflight)
    forged = replace(inventory, inventory_sha256="f" * 64)
    with pytest.raises(PendingDecisionInventoryError, match="digest mismatch"):
        validate_pending_decision_inventory(forged)
