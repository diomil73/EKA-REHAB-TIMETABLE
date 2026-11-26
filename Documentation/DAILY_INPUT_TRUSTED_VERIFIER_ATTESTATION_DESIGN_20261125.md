# DIOMIL Trusted Verifier Attestation: Threat Model and Deployment Decisions
Date: 2026-11-25
Scope: DAILY_INPUT phase verifier reports and transaction receipts
Status: PRE-AUTHORIZATION DESIGN; NO PRODUCTION TRUST ROOT, NO SIGNER, NO APPROVAL

## A. Security objective

Distinguish three independent layers:
1. **Consistency**: source/ledger/archive facts satisfy independent readback rules.
2. **Integrity**: a report, evidence artifact, and receipt chain match pinned bytes and operation identity.
3. **Authenticity/provenance**: a trusted, independently administered verifier actually inspected those bytes with the intended implementation and security configuration.

The existing phase report and receipt codecs implement some of layer 2. A SHA-256 checksum, Git commit ID, declared `verifier_id`, recorded Excel version, or JSON `VERIFIED` value is not a cryptographic claim of layer 3. A writer, UI macro, operator-supplied file or verifier executable cannot authenticate itself.

## B. Threats the mechanism must resist

- Forged VERIFIED reports with recomputed self-digests.
- Replaying old genuine evidence for another operation, date, source hash, phase, revision, or workbook.
- Replacing the verifier executable, changing its approved build, or downgrading its version.
- Switching evidence payload or Excel workbook between inspection and signature.
- Self-signed manifests, attacker-supplied trust roots, an embedded private key in workbook/VBA/Python or a secret committed to Git.
- Reordered or incomplete multi-kind transition evidence sets, duplicate receipts, replay across revisions.
- Expired, revoked, rotated or compromised signing keys and ambiguous signing timestamps.
- A crash after reading but before durable evidence/receipt/journal publication.
- Insider/operator modification of both report and digest.

## C. Required trust architecture, not implemented

An independently administered **trust root** must pin approved public verification keys outside writable workbook and report artifacts. Only an administrator-approved provider may sign canonical report claims after verifying the expected build identity, measured executable hash, exact workbook/evidence SHA-256, operation/revision and verifier execution conditions. A claimed `verifier_build_sha256` is untrusted unless checked by the designated provider.

Candidate custody options requiring a deployment decision:
- Windows CNG/TPM-backed non-exportable signing key with ACLs restricting signing to a separately installed verifier service.
- Independently managed signing service with hardware-backed key custody, auditable access and availability/crash policy.

Do not make a choice until operational constraints, Windows deployment permissions, certificate lifecycle and the user's environment have been reviewed. Never generate or store private keys in Git, XLSM, PowerShell snippets, Python source or test fixtures. The only repository fixture signature may be a deliberately invalid 64-byte placeholder in negative tests. It is never a cryptographic proof.

If Ed25519 is selected for report claims, the actual cryptographic signature verifier must be implemented and tested using an authenticated, externally pinned public key and signature library/facility. The current parser supports only the syntax `Ed25519` but does **not** verify it. No `signature_valid` Boolean supplied by the caller may count as verification.

## D. Canonical claim proposal (non-authorizing code contract)

`Python/rehab_excel/daily_input_verifier_attestation.py` defines `UntrustedVerifierAttestationClaim` and `decode_untrusted_attestation_claim()`. It requires exact canonical UTF-8 JSON fields:

- schema, verifier_id, contract_version, evidence_kind
- report_sha256, verifier_build_sha256, source_workbook_sha256
- operation_id, service_date, journal_revision, journal_phase
- signer_key_id, signature_algorithm, signature_b64
- issued_at_utc, expires_at_utc

The claim is not a verdict. Merely checking canonical JSON, field syntax, Ed25519-shaped signature bytes and report identity cannot authenticate the signer. Accordingly `require_trusted_verifier_attestation()` **always raises** in the present implementation, including when all fields match.

For a future authorized implementation, the signing payload must be explicitly versioned and domain-separated, and the signature must cover all claim fields *except* the signature itself. It must also bind the **exact report self-digest**, which in turn binds evidence artifact bytes and inspection scope. The verifier registry/build manifest must be pinned independently; user-provided `signer_key_id` is an identifier for lookup, not a trusted key. Identity, issuance/expiration windows, revocation and monotonic/replay protection must be evaluated using a trusted source of time and durable state.

## E. Mandatory authorization order (future work)

1. Validate original source, preflight facts, operation transition, required evidence kinds.
2. Independently verify saved OOXML and native Excel COM views for all required scopes.
3. Hash and retain the exact evidence artifact, with durability and path confinement.
4. Validate each phase report and its self-digest.
5. Verify build manifest and attestation using independently pinned trust root and policy.
6. Validate the complete ordered evidence bundle and every receipt artifact/chain.
7. Persist authenticated receipts durably in a location not writable by the workbook UI.
8. Only then consider guarded journal append, independent reopen and crash recovery.
9. Never let verification results directly bypass the existing Gate 2 operational acceptance.

## F. Open administrator decisions and blockers

- Trust root/key custody and secure installation/update mechanism.
- Approved verifier binary hash and release signing process.
- Who may invoke verifier and who may sign; principle of separation of duties.
- Offline mode, key rotation and revocation; trusted clock and replay inventory.
- Signature algorithm/provider availability on target Windows Excel environment.
- Retention and recovery of evidence and signed receipts, privacy/PHI minimization.
- Exact source and operational definition of a valid zero-change day.

All are BLOCKED for positive authorization until a deliberate operational design is approved.

## G. Current tests and explicit limits

Synthetic negative tests cover canonical claim parsing, invalid algorithm and signature shape, bad timestamps, unexpected fields, report/source/operation drift, tampered report self-digest and unconditional denial of otherwise matching self-asserted claims.

They do not execute a signer, verify any real cryptographic signature, prove a trusted time/key custody facility, run native Excel COM, alter operational workbooks, or authorize journal Apply/Restore. PREPARED_SNAPSHOT, INPUT_AND_LEDGER_REOPEN and PROJECTIONS_REOPEN are REPORT_ADAPTER_ONLY; the other seven verifier kinds remain NOT_IMPLEMENTED. None of these report adapters is trusted or authorizing. Gate 2 stays closed.


### CI confirmation for attestation claim contract — 25/11/2026

GitHub Actions workflow `45598412259` for code head `a66088e91792b49095559bccdf9c800ede2ad7b0` completed SUCCESS:
- Windows / Python 3.11: **2756 passed**, 68 warnings, 44.39s.
- Ubuntu / Python 3.11: **2756 passed**, 68 warnings, 30.26s.
- Both synthetic safety jobs: success.

The workflow uses `--disable-warnings`; warning details are suppressed but warning totals remain visible. No native COM was run. These are tests of fail-closed metadata handling, not proof of cryptographic attestation or operational Apply/Restore. The trusted verifier signer/key custody/build pinning remain unimplemented; all claims are rejected by the current attestation boundary. Gate 2 stays closed.

### 2026-11-26 — PROJECTIONS_REOPEN report-only adapter added

The first three evidence kinds now have report-only adapters: PREPARED_SNAPSHOT, INPUT_AND_LEDGER_REOPEN and PROJECTIONS_REOPEN. The projection adapter composes native source capture, input/ledger, MASTER and full therapist-grid readbacks, then publishes an operation/date/source-bound report. Its synthetic tests do not authenticate the verifier or execute a user's real Excel workbook.

The current registry therefore contains **3 REPORT_ADAPTER_ONLY** and **7 NOT_IMPLEMENTED** kinds. No kind is `IMPLEMENTED`; trusted signer/build provenance remains unresolved and all positive report validation remains fail-closed. The report distinguishes original `operation.source_sha256` from the current post-commit `evidence_sha256`. Gate 2 remains closed.
