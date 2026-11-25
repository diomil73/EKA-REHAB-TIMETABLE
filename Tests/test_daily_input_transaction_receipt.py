from dataclasses import replace
from hashlib import sha256

import pytest

from rehab_excel.daily_input_operation import DailyInputOperation, DailyInputOperationPhase as Phase
from rehab_excel.daily_input_phase_evidence import EvidenceKind
from rehab_excel.daily_input_transaction_receipt import (
    DailyInputTransactionReceipt,
    TransactionReceiptError,
    create_transaction_receipt,
    decode_transaction_receipt,
    validate_receipt_binding,
    verify_receipt_artifact_bytes,
)


def previous_operation():
    return DailyInputOperation("op-20261125-001", "a" * 64, "2026-11-25")


def operation():
    return DailyInputOperation(
        "op-20261125-001", "a" * 64, "2026-11-25",
        phase=Phase.INPUT_AND_LEDGER_COMMITTED, revision=1,
    )


def receipt_fields(**overrides):
    value = {
        "schema": "diomil-transaction-receipt-v1",
        "receipt_id": "receipt-001",
        "operation_id": "op-20261125-001",
        "source_workbook_sha256": "a" * 64,
        "service_date": "2026-11-25",
        "journal_revision": 0,
        "journal_phase": "PREPARED",
        "evidence_kind": EvidenceKind.PREPARED_SNAPSHOT.value,
        "evidence_artifact": "op-20261125-001/prepared.bin",
        "evidence_sha256": sha256(b"fictional evidence").hexdigest(),
        "verifier_id": "native-reopen-verifier",
        "verifier_version": "v1",
        "verification_method": "native_excel_readonly_reopen",
        "verification_result": "VERIFIED",
        "verified_at_utc": "2026-11-25T10:30:00Z",
        "previous_receipt_sha256": None,
    }
    value.update(overrides)
    return value


def input_ledger_receipt_fields(**overrides):
    """A later-phase receipt fixture linked to its PREPARED predecessor."""
    value = receipt_fields(
        journal_revision=1,
        journal_phase=Phase.INPUT_AND_LEDGER_COMMITTED.value,
        evidence_kind=EvidenceKind.INPUT_AND_LEDGER_REOPEN.value,
        previous_receipt_sha256="b" * 64,
    )
    value.update(overrides)
    return value


def test_receipt_round_trip_is_canonical_and_digest_checked():
    receipt = create_transaction_receipt(**receipt_fields())
    encoded = receipt.to_json()
    decoded = decode_transaction_receipt(encoded)
    assert decoded == receipt
    assert decode_transaction_receipt(encoded.encode("utf-8")) == receipt


def test_receipt_digest_detects_tampering():
    receipt = create_transaction_receipt(**receipt_fields())
    changed = receipt.to_dict()
    changed["verifier_version"] = "v2"
    with pytest.raises(TransactionReceiptError, match="SHA-256 mismatch"):
        decode_transaction_receipt(changed)


def test_receipt_rejects_missing_and_extra_fields():
    fields = receipt_fields()
    fields.pop("verifier_id")
    with pytest.raises(TransactionReceiptError, match="missing or unexpected"):
        create_transaction_receipt(**fields)
    fields = receipt_fields(unexpected_field="nope")
    with pytest.raises(TransactionReceiptError, match="missing or unexpected"):
        create_transaction_receipt(**fields)


@pytest.mark.parametrize("field,value", [
    ("source_workbook_sha256", "A" * 64),
    ("evidence_sha256", "bad"),
    ("journal_revision", True),
    ("service_date", "25-11-2026"),
    ("verified_at_utc", "2026-11-25T10:30:00"),
    ("verification_result", "MAYBE"),
    ("evidence_kind", "CALLER_SAYS_OK"),
])
def test_receipt_rejects_malformed_fields(field, value):
    with pytest.raises(TransactionReceiptError):
        create_transaction_receipt(**receipt_fields(**{field: value}))


@pytest.mark.parametrize("locator", [
    "../outside.bin",
    "op/../../outside.bin",
    "/absolute/path.bin",
    "C:/outside.bin",
    "op\\outside.bin",
    "op//file.bin",
])
def test_receipt_rejects_unsafe_artifact_locator(locator):
    with pytest.raises(TransactionReceiptError, match="evidence_artifact"):
        create_transaction_receipt(**receipt_fields(evidence_artifact=locator))


def test_revision_zero_requires_prepared_and_no_predecessor():
    with pytest.raises(TransactionReceiptError, match="Revision zero"):
        create_transaction_receipt(**receipt_fields(journal_phase="COMPLETE"))
    with pytest.raises(TransactionReceiptError, match="Revision zero"):
        create_transaction_receipt(**receipt_fields(previous_receipt_sha256="b" * 64))


def test_prepared_snapshot_evidence_is_exclusive_to_revision_zero():
    with pytest.raises(TransactionReceiptError, match="PREPARED_SNAPSHOT"):
        create_transaction_receipt(**receipt_fields(
            evidence_kind=EvidenceKind.FAILURE_DIAGNOSTIC.value,
        ))
    with pytest.raises(TransactionReceiptError, match="PREPARED_SNAPSHOT"):
        create_transaction_receipt(**receipt_fields(
            journal_revision=1,
            journal_phase=Phase.INPUT_AND_LEDGER_COMMITTED.value,
            evidence_kind=EvidenceKind.PREPARED_SNAPSHOT.value,
            previous_receipt_sha256="b" * 64,
        ))


def test_nonzero_revision_requires_predecessor_digest():
    with pytest.raises(TransactionReceiptError, match="requires previous"):
        create_transaction_receipt(**receipt_fields(
            journal_revision=1, journal_phase=Phase.INPUT_AND_LEDGER_COMMITTED.value,
        evidence_kind=EvidenceKind.INPUT_AND_LEDGER_REOPEN.value,
        ))


def test_binding_matches_operation_phase_evidence_and_pinned_verifier():
    op = operation()
    receipt = create_transaction_receipt(**input_ledger_receipt_fields())
    validate_receipt_binding(
        receipt, op, previous_operation=previous_operation(),
        expected_previous_receipt_sha256="b" * 64,
        verifier_registry={"native-reopen-verifier": "v1"},
    )


def test_binding_rejects_operation_or_evidence_drift():
    op = operation()
    receipt = create_transaction_receipt(**input_ledger_receipt_fields())
    with pytest.raises(TransactionReceiptError, match="identity/revision"):
        validate_receipt_binding(
            receipt, replace(op, source_sha256="b" * 64),
            previous_operation=previous_operation(),
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry={"native-reopen-verifier": "v1"},
        )
    with pytest.raises(TransactionReceiptError, match="evidence kind"):
        validate_receipt_binding(
            receipt, op, previous_operation=previous_operation(),
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry={"native-reopen-verifier": "v1"},
        )


def test_binding_rejects_unpinned_verifier_and_wrong_predecessor():
    receipt = create_transaction_receipt(**input_ledger_receipt_fields())
    args = dict(
        previous_operation=previous_operation(),
        expected_previous_receipt_sha256="b" * 64,
        verifier_registry={"native-reopen-verifier": "v2"},
    )
    with pytest.raises(TransactionReceiptError, match="not pinned"):
        validate_receipt_binding(receipt, operation(), **args)
    with pytest.raises(TransactionReceiptError, match="predecessor"):
        validate_receipt_binding(
            receipt, operation(), previous_operation=previous_operation(),
            expected_previous_receipt_sha256="c" * 64,
            verifier_registry={"native-reopen-verifier": "v1"},
        )


def test_artifact_verifier_checks_exact_bytes_and_root_containment(tmp_path):
    root = tmp_path / "evidence"
    target = root / "op-20261125-001" / "prepared.bin"
    target.parent.mkdir(parents=True)
    target.write_bytes(b"fictional evidence")
    receipt = create_transaction_receipt(**receipt_fields())
    assert verify_receipt_artifact_bytes(receipt, root) == receipt.evidence_sha256
    target.write_bytes(b"changed")
    with pytest.raises(TransactionReceiptError, match="SHA-256 mismatch"):
        verify_receipt_artifact_bytes(receipt, root)


def test_artifact_verifier_rejects_symlink_escape(tmp_path):
    root = tmp_path / "evidence"
    root.mkdir()
    outside = tmp_path / "outside.bin"
    outside.write_bytes(b"fictional evidence")
    link = root / "escape.bin"
    try:
        link.symlink_to(outside)
    except (OSError, NotImplementedError):
        pytest.skip("Symlinks unavailable on this platform")
    receipt = create_transaction_receipt(**receipt_fields(evidence_artifact="escape.bin"))
    with pytest.raises(TransactionReceiptError, match="escapes evidence root"):
        verify_receipt_artifact_bytes(receipt, root)


def test_receipt_is_not_an_authorization_surface():
    import inspect
    from rehab_excel import daily_input_transaction_receipt as module
    source = inspect.getsource(module)
    assert "does NOT authenticate" in source
    assert "DailyInputJournalApplyEnabled = True" not in source
    assert "ThisWorkbook.Save" not in source
    assert not hasattr(module, "enable_apply")

def test_binding_rechecks_digest_even_for_direct_dataclass_instances():
    receipt = create_transaction_receipt(**input_ledger_receipt_fields())
    forged = replace(receipt, verifier_version="v2")
    with pytest.raises(TransactionReceiptError, match="SHA-256 mismatch"):
        validate_receipt_binding(
            forged, operation(),
            previous_operation=previous_operation(),
            expected_previous_receipt_sha256="b" * 64,
            verifier_registry={"native-reopen-verifier": "v2"},
        )
