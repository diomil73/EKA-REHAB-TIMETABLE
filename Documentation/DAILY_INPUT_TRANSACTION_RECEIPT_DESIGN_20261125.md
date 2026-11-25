# DAILY_INPUT Transaction Receipt System — Design Contract
Date: 2026-11-25
Status: DESIGN ONLY — NOT IMPLEMENTED, NOT AUTHORIZING, NOT OPERATIONAL

## 1. Purpose

Define the durable evidence envelope needed to bind one DAILY_INPUT journal revision to the exact operation and source workbook; exact phase and journal revision; saved workbook/evidence artifact actually reopened; independent verifier; verifier result and evidence digest; and prior receipt in the same operation chain.

This document does not enable Apply, Restore, Replacement, Amendment, CloseDay or Gate 2. It does not claim that any receipt exists.

## 2. Existing source contracts to preserve

- Journal remains append-only: OperationID, Revision, Phase, PayloadSHA256, Payload.
- Journal payload binds operation_id, source_sha256, service_date, phase, revision, failure_code, failure_detail, operator_resolution and rollback_verified.
- Pure operation state machine remains authority for legal phase transitions.
- PREPARED native VBA bridge is diagnostic-only. DailyInputJournalApplyEnabled() remains False.
- Phase evidence policy is a list of obligations, not evidence or authorization.
- Full rollback auditor is read-only and requires prepared-snapshot digest to be bound by a separately verified durable receipt.
- SHA-256 detects changes; it does not prove who created evidence or that a verifier was trustworthy.

## 3. Proposed v1 receipt envelope

Canonical UTF-8 JSON with exactly these fields:

| Field | Requirement |
|---|---|
| schema | Fixed: diomil-transaction-receipt-v1 |
| receipt_id | Unique, non-empty identifier |
| operation_id | Exact DAILY_INPUT operation ID |
| source_workbook_sha256 | 64 lowercase hex; equals journal source_sha256 |
| service_date | Exact operation service date |
| journal_revision | Non-negative integer matching target journal revision |
| journal_phase | Exact phase enum value for that revision |
| evidence_kind | Exact EvidenceKind required for the phase |
| evidence_artifact | Relative operation-scoped locator; absolute paths and traversal forbidden |
| evidence_sha256 | SHA-256 of exact saved evidence artifact bytes |
| verifier_id | Stable independent verifier implementation identifier |
| verifier_version | Immutable verifier build/version |
| verification_method | Enumerated method; native Excel reopen for workbook evidence |
| verification_result | Fixed VERIFIED; failures produce diagnostic, never success receipt |
| verified_at_utc | RFC3339 UTC timestamp |
| previous_receipt_sha256 | Previous receipt digest; null only for PREPARED |
| receipt_sha256 | SHA-256 of canonical JSON excluding this field |

No caller-supplied Boolean can satisfy a proof obligation. No receipt may refer to a workbook path alone: exact bytes must be hashed. No absolute local path is stored; locator is relative to a controlled evidence root.

## 4. Binding and validation rules

1. Receipt operation_id, source hash, service date, journal revision and phase must match the independently decoded append-only journal row.
2. evidence_kind must be one of required_evidence_for_transition(previous, candidate).
3. Evidence artifact must exist under configured evidence root, resolve without path traversal or symlink escape, and hash to evidence_sha256.
4. Workbook evidence must be produced by an independent native Excel read-only reopen verifier over saved bytes. A Python model, in-memory workbook object, VBA return value or caller assertion is insufficient.
5. Verifier identity/version must be pinned by application build manifest. A caller-supplied string is not verifier authentication.
6. Receipt chain ordering must match journal revision ordering. No phase skips or cross-operation/source links.
7. PREPARED snapshot evidence must be captured before first business mutation and bound to PREPARED revision.
8. COMPLETE requires full-workbook final native reopen evidence. ROLLED_BACK requires strict full-workbook rollback audit plus independent native reopen and a receipt binding exact PREPARED snapshot digest.
9. Receipt digest is an integrity checksum, not a signature. It cannot independently prove origin or authenticity.
10. No phase may be appended as verified until receipt and all required evidence artifacts have been independently checked and durably retained.

## 5. Trust boundary — unresolved technical prerequisite

The repository has no demonstrated trusted signing/attestation mechanism for receipts. JSON plus SHA-256 chain is tamper-evident only if storage is trusted; a party able to rewrite both receipt and digest can forge it.

Therefore v1 must remain non-authorizing until a protected local signing key, OS-backed signing/attestation facility, or independently controlled receipt service is designed and tested. Do not invent a secret, embed a key in VBA/Python source, or treat Git commit identity as a runtime transaction signer. User/administrator decision is required only when choosing production trust root and key custody. Until then, implement and test structural receipt validation only.

## 6. Required independent verifier registry

A verifier registry must map each EvidenceKind to one pinned verifier and strict output schema:
- INPUT_AND_LEDGER_REOPEN: native reopen of DAILY_INPUT and authoritative ledger.
- PROJECTIONS_REOPEN: native reopen of MASTER_SCHEDULE and all affected therapist projections.
- PENDING_DECISION_INVENTORY: native reopen of pending replacement inventory.
- REPLACEMENTS_REOPEN: native reopen of replacement facts and suspensions.
- OPERATOR_RESOLUTION: independent verification of operator decision and operation linkage.
- APPLIED_ARCHIVE_REOPEN: native reopen of applied archive.
- FINAL_WORKBOOK_REOPEN: native reopen of final saved workbook and required invariants.
- FAILURE_DIAGNOSTIC: durable operation-bound failure diagnostic.
- FULL_WORKBOOK_ROLLBACK: strict package audit plus native reopen of restored workbook and PREPARED snapshot identity.

A verifier must return a structured report derived from inspected bytes. It must not accept caller-provided verified=true.

## 7. Persistence order and crash recovery

For each phase:
1. Capture and retain required immutable evidence artifacts.
2. Close writer and reopen exact saved bytes using designated independent verifier.
3. Validate verifier output, operation/source identity, artifact hashes and phase requirements.
4. Create canonical receipt and validate schema, digest and chain.
5. Persist receipt durably.
6. Append journal phase revision.
7. Reopen and verify full append-only journal history.
8. Retain artifacts until COMPLETE or verified ROLLED_BACK, plus configured audit-retention period.

Crash recovery must distinguish: evidence exists but receipt absent; receipt exists but journal append absent; journal append exists but readback incomplete; journal points to phase but evidence is missing/changed. Every ambiguous state fails closed. Recovery must be idempotent and never infer success from an incomplete pair.

## 8. Rollback receipt specifics

Rollback chain must bind PREPARED snapshot relative locator and exact SHA-256; source hash and operation ID; failure diagnostic and failed journal revision; restored workbook exact SHA-256; strict non-journal package comparison; journal prefix and terminal ROLLED_BACK revision; independent native Excel reopen report; verifier build identity and verification time.

Existing full rollback auditor remains read-only. A separate rollback executor restores bytes; auditor independently verifies them. Executor cannot self-certify success.

## 9. Minimum tests before any native positive transaction

- Canonical JSON determinism and digest verification.
- Reject missing, extra, mistyped and malformed fields.
- Reject invalid hashes, timestamps, phase/evidence mismatch and path traversal.
- Reject operation/source/service-date/revision drift against journal.
- Reject chain gaps, duplicate IDs, reordered receipts and changed prior digest.
- Reject evidence artifact byte changes after receipt creation.
- Reject caller-supplied success flags and unknown verifier IDs/versions.
- Reject unsigned receipts at any authorization boundary.
- Simulate every crash window in section 7 and verify fail-closed idempotent recovery.
- Verify receipt module changes no workbook and enables no Gate 2.

Initial tests use fictional synthetic artifacts. Native Excel tests use disposable synthetic XLSM copies. Original CC remains untouched.

## 10. Explicit non-goals

The new Python module implements only structural receipt encoding/decoding, operation/evidence/verifier-label binding, and evidence-file byte/path verification. It does not authenticate receipt origin. This receipt design does not implement transaction orchestrator, native phase verifiers, rollback executor, VBA hooks, UI controls or operational Apply/Restore. It does not prove positive transactions. Statistics, dashboards and management analytics remain outside first practice DEMO critical path.

## 11. Current readiness

- Journal codec: implemented as pure append-only codec.
- Lifecycle state model: implemented as pure model.
- PREPARED native bridge: implemented, diagnostic-only.
- Independent phase verifiers: not registered.
- Structural receipt codec/validator: implemented in daily_input_transaction_receipt.py with synthetic tests added; tests not yet executed in this session.
- Trusted receipt writer/authenticator: not implemented.
- Trusted signing/attestation root: not selected or implemented.
- Phase orchestrator: not implemented.
- Full rollback executor: not proven.
- Gate 2: remains closed.
## Addendum — PREPARED snapshot evidence is distinct

The initial PREPARED revision (revision 0) must carry EvidenceKind.PREPARED_SNAPSHOT = prepared_snapshot_native_readback. FAILURE_DIAGNOSTIC is reserved for the FAILED phase and is not valid PREPARED evidence. PREPARED_SNAPSHOT is not valid for any later revision. The structural receipt codec enforces this phase/revision distinction. An existing PREPARED native readback diagnostic is documented in the continuity guide §4.144, but it is not yet adapted to emit this phase-verifier report and receipt chain. Trusted attestation remains unimplemented.

## 10. Phase-verifier contract registry and report envelope — contract only, no verifier implementation — 25/11/2026

Added:
- Python/rehab_excel/daily_input_phase_verifier_contract.py
- Tests/test_daily_input_phase_verifier_contract.py

The new module defines one immutable contract for each EvidenceKind and a strict canonical JSON report envelope. Contracts name the verifier ID, contract version, verification method, exact inspection scope, and whether native Excel reopen, package audit, and PREPARED snapshot binding are required. The report binds operation ID, source workbook SHA-256, service date, journal revision/phase, evidence kind, verifier contract identity, method, result, relative artifact locator/hash, inspected scope, native Excel version where applicable, package audit digest where required, PREPARED snapshot digest where required, UTC timestamp and report digest.

Evidence mapping follows this document's §6:
- PREPARED_SNAPSHOT: saved workbook bytes, PREPARED snapshot identity and PREPARED journal revision.
- INPUT_AND_LEDGER_REOPEN: DAILY_INPUT plus authoritative ledger.
- PROJECTIONS_REOPEN: MASTER_SCHEDULE plus affected therapist projections and source linkage.
- PENDING_DECISION_INVENTORY: pending replacement inventory.
- REPLACEMENTS_REOPEN: replacement facts and suspensions.
- OPERATOR_RESOLUTION: independent operation-bound decision-record audit.
- APPLIED_ARCHIVE_REOPEN: applied archive.
- FINAL_WORKBOOK_REOPEN: full workbook and required invariants.
- FAILURE_DIAGNOSTIC: retained operation-bound failure artifact.
- FULL_WORKBOOK_ROLLBACK: restored workbook, non-journal package comparison, journal prefix and exact PREPARED snapshot identity.

Critical status: PREPARED_SNAPSHOT is EXISTING_DIAGNOSTIC_ONLY because verify_daily_input_journal_prepared.py already supports independent saved-byte inspection and optional native Excel read-only reopen; it is not yet integrated with this report/receipt contract. The remaining nine contracts are NOT_IMPLEMENTED, and none is IMPLEMENTED for positive phase authorization. The registry is immutable metadata, not an executable verifier registry or trust root. The report codec validates schema, canonical digest, operation/phase binding and contract scope; it does not execute Excel, authenticate a report, or prove that the claimed inspection occurred. A synthetic report marked VERIFIED is still rejected for positive phase use while its contract is NOT_IMPLEMENTED or only EXISTING_DIAGNOSTIC_ONLY. No caller-supplied verified Boolean is accepted. The required evidence list is derived internally from required_evidence_for_transition(previous, candidate), not passed by the caller.

The module remains non-authorizing. SHA-256 is not a signature. No journal append, Apply enablement, VBA write, workbook mutation or positive transaction path is exposed. The next implementation step is to adapt the already native-proven PREPARED readback diagnostic into the report/receipt envelope without repeating its completed native diagnostic; then implement the independent INPUT_AND_LEDGER_REOPEN native Excel verifier, plus separately verifiable non-Excel decision/failure artifact auditors. Trusted build-manifest/attestation provenance remains unresolved. Tests have been added but NOT YET EXECUTED locally.
