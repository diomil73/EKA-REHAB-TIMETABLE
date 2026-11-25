# DAILY_INPUT Internal Verifier Review — Offline Source Audit
Date: 2026-11-25
Reviewed branch: `work/afternoon-medical-scope-20261115`
Reviewed HEAD before this documentation commit: `817941256d47581dbba51820a4077a9cec72b043`
Status: SOURCE REVIEW ONLY — NO LOCAL TEST EXECUTION — NO EXCEL/COM EXECUTION — NON-AUTHORIZING

## 1. Scope and method

Reviewed the current source for:
- `daily_input_phase_evidence.py`
- `daily_input_phase_verifier_contract.py`
- `daily_input_transaction_receipt.py`
- `daily_input_operation.py`
- `daily_input_applied_state.py`
- `daily_input_prepared_phase_adapter.py`
- `verify_daily_input_journal_prepared.py`
- `daily_input_applied_archive.py`
- `daily_input_full_rollback_audit.py`
- `daily_replacement_facts.py`
- `daily_replacement_fact_batch.py`
- Existing focused tests and the continuity / delivery / test-baseline documents.

This review used repository source only. It did not open Excel, execute Python, validate a real XLSM, invoke COM, or change operational workbook/VBA handlers.

## 2. Current evidence boundary

The ten evidence kinds and their lifecycle obligations are present in the immutable contract registry. PREPARED has a report-producing adapter over the existing native readback and is `REPORT_ADAPTER_ONLY`. The other nine evidence kinds remain `NOT_IMPLEMENTED`. No verifier is currently `IMPLEMENTED` for positive phase authorization.

The existing `verify_input_ledger_saved_readback()` is a useful independent OOXML consistency check. It checks the saved DAILY_INPUT against separately supplied preflight changes, checks the active DAILY_CHANGE_LEDGER membership and veryHidden state, and verifies that the workbook bytes did not change during the readback. Its own contract explicitly says it is not a native Excel reopen or authorization. It currently requires a non-empty `validated_changes` tuple; the native verifier design must preserve this behavior unless a separately reviewed business rule defines the valid zero-change case.

The existing PREPARED native COM reader is a useful pattern: fresh Excel instance, macros force-disabled, events/alerts/link updates disabled, read-only open, explicit worksheet protection/visibility checks, exact row comparison, close without saving and quit. Reuse the safety pattern, not its PREPARED-specific data reader.

## 3. Findings that must be closed before positive phase authorization

### F-01 — Multi-obligation transition completeness is not yet enforced as one atomic evidence set

`required_evidence_for_transition()` can return more than one obligation. In particular, the transition from `PENDING_OPERATOR_DECISION` to `REPLACEMENTS_RESOLVED` requires both `OPERATOR_RESOLUTION` and `REPLACEMENTS_REOPEN`.

The phase-report validator validates one report at a time and confirms that its evidence kind is a member of the required tuple. The receipt validator likewise validates one receipt at a time and confirms membership. Neither API, by itself, proves that every required evidence kind for that transition has been supplied exactly once.

Required design correction before any positive integration: introduce a transition-evidence bundle validator. It must derive the exact required evidence set internally, require one valid report and receipt for every required kind, reject missing/extra/duplicate kinds, bind every item to the same operation/source/date/revision/phase, validate each artifact's exact bytes, and verify a single unambiguous predecessor/chain order. The journal transition must remain impossible until the complete bundle validates. Do not treat one successful report as completion of a multi-obligation transition.

### F-02 — Receipt binding currently accepts a caller-provided evidence-obligation tuple

`validate_receipt_binding()` accepts `required_evidence` from its caller. This is weaker than `validate_phase_verifier_report()`, which derives obligations from the previous and candidate operations using the lifecycle policy.

Before positive use, receipt validation must receive the previous operation (or a separately validated transition object) and derive the required evidence internally. Do not allow a caller to choose a narrower evidence list. The PREPARED revision-zero case must remain a separate, explicit initial-state rule.

### F-03 — Receipt/report integrity is not authenticity

Canonical JSON and SHA-256 detect corruption and bind bytes; they do not prove who created a report, which verifier binary ran, or whether its output is trustworthy. The receipt verifier registry is currently a caller-supplied mapping. No trusted signing/attestation root or trusted build-manifest verification is present.

Keep all adapters non-authorizing. Before enabling any phase append, define a trusted verifier identity/version source and authenticated report/receipt provenance. The production trust-root and key-custody choice remains an administrator decision; do not invent, embed, or silently generate a signing key in VBA/Python.

### F-04 — INPUT_AND_LEDGER needs two distinct proofs, not a relabelled OOXML check

The existing OOXML readback must remain intact and be recorded as its own consistency result. The new native verifier must open the exact saved XLSM in a fresh native Excel COM instance, read-only, with macros/events/alerts/link updates disabled, and independently inspect the DAILY_INPUT and DAILY_CHANGE_LEDGER values as Excel exposes them.

The COM proof must bind to the same operation, service date, preflight facts and exact workbook bytes. It must compare complete change identities, not counts alone; check ledger schema, active/restored semantics, uniqueness, protection/visibility requirements and reject formula/error cells in the authoritative ledger region. It must close without saving, quit Excel, and re-hash workbook and payload/evidence before report publication.

Do not claim that native COM merely reopening the file proves the OOXML reader's facts. Both views must be independently reconciled to the same trusted preflight facts. Resolve explicitly how the product treats an empty validated change set before implementing this phase; current OOXML API rejects it.

### F-05 — Downstream verifier scopes must remain phase-specific

The registry scopes are obligations, not implementation evidence. Preserve these boundaries when each verifier is built:

| Evidence kind | Required independent scope |
|---|---|
| PROJECTIONS_REOPEN | MASTER_SCHEDULE plus every affected therapist projection, linked to the exact committed input/ledger facts |
| PENDING_DECISION_INVENTORY | Complete pending replacement inventory for the operation/date, not a partial queue |
| REPLACEMENTS_REOPEN | Replacement facts and suspensions, exact target/session/provider identities, status, uniqueness and operation linkage |
| OPERATOR_RESOLUTION | Separate durable operator decision record, actor/choice/time/reason and operation linkage; this is a non-Excel evidence class |
| APPLIED_ARCHIVE_REOPEN | Exact applied archive state for all relevant dates, typed payload/digests, protected veryHidden storage and consistency with the applied grid/ledger |
| FINAL_WORKBOOK_REOPEN | Full required workbook invariants and operation identity, not merely successful open or a subset of sheets |
| FAILURE_DIAGNOSTIC | Retained, operation-bound failure artifact with exact failure code/detail and artifact hash; no success receipt for a failed verification |
| FULL_WORKBOOK_ROLLBACK | PREPARED snapshot binding, restored business state, exact non-journal package comparison, journal history and final native Excel read-only reopen |

The existing full rollback auditor deliberately compares every non-journal OOXML package member byte-for-byte and can reject a semantically correct Excel save if Excel rewrites unrelated package parts. Do not weaken this comparison by ignoring unexplained differences. If a future semantic fallback is needed, it must be a separately specified and independently evidenced proof, not a relaxation of the current audit.

### F-06 — Native Excel lifecycle and resource discipline are part of every verifier contract

Every COM verifier must use a fresh isolated Excel instance; set force-disable macro security before opening; disable events, alerts and link updates; open ReadOnly; verify ReadOnly; never Save/SaveAs; close with SaveChanges=False; always Quit in `finally`; and fail closed if cleanup or readback is incomplete. Record Excel version. Compare file hashes before and after the full native readback. Tests must prove the cleanup path on both success and injected failures.

Do not reuse a live user's Excel process or an already-open workbook as verification evidence.

## 4. Ordered implementation plan

1. Finish the PREPARED local focused rerun and warning investigation. The latest user-provided run (before commits `1f4a6c7` and `8179412`) was 109 passed, 1 skipped, 17 warnings. The two follow-up `keep_vba=False` commits have not yet been locally executed. No warning-removal claim is made.
2. Implement INPUT_AND_LEDGER native COM read-only verifier and report adapter, reusing the existing OOXML comparator rather than duplicating it. Keep status `REPORT_ADAPTER_ONLY`.
3. Add the transition evidence bundle validator and remove caller-selected receipt evidence obligations before any positive phase acceptance.
4. Implement projections verifier.
5. Implement pending-decision inventory and replacement facts/suspensions verifiers.
6. Implement the non-Excel operator-resolution record auditor.
7. Implement applied archive verifier.
8. Implement full-workbook final verifier.
9. Implement failure-artifact auditor.
10. Complete full rollback proof using the retained PREPARED snapshot, strict package audit, native Excel read-only reopen and authenticated evidence chain.
11. Only after a separately trusted attestation mechanism and all phase verifiers exist, design the transaction orchestrator and test crash recovery/idempotency. Gate 2 remains closed until a separate operational acceptance decision.

## 5. Test matrix required for the next INPUT_AND_LEDGER verifier

Synthetic tests should cover:
- COM exact-match success and native Excel version capture.
- Macros forced disabled, events/alerts/link updates disabled, ReadOnly open and no Save.
- Workbook.Close(SaveChanges=False) and Excel.Quit on success and every failure branch.
- COM/OOXML disagreement for date, change ID, type, subject, display, reason, source row and status.
- Missing/wrong/visible/unprotected ledger; wrong headers; formula/error cells; duplicate IDs; duplicate active target facts; unknown statuses.
- Empty changes policy explicitly, without silently treating missing evidence as an empty day.
- Workbook or payload mutation during verification.
- Evidence-root escape, symlink escape, report overwrite and report tampering.
- PREPARED snapshot and operation identity mismatch.
- Rejected/indeterminate results cannot create a success receipt.
- Adapter report remains non-authorizing while trust attestation is absent.

Use injected fake COM objects for deterministic tests. A separate disposable real-XLSM/native-Excel smoke test is still required when the PC is available. Synthetic COM tests are not proof that the user's actual Excel installation behaves identically.

## 6. Work that can proceed without the user's PC

- Source and contract reviews in GitHub.
- Cross-phase evidence completeness and crash-window analysis.
- Design and synthetic test specifications.
- Documentation and coverage-register updates.
- Review of existing handlers against phase evidence obligations.

## 7. Work that must wait for the PC

- Local pytest rerun after the latest warning-related commits.
- Any native Excel COM execution against a disposable XLSM.
- Actual Excel version/behavior confirmation.
- Real workbook Save/Close/Reopen smoke tests.
- Operational acceptance and any positive transaction test.

No production workbook, VBA, UI, Apply/Restore path, journal writer or Gate 2 setting was changed by this review.


## 8. Transition evidence bundle validator implemented — 2026-11-25

Added:
- `Python/rehab_excel/daily_input_transition_evidence.py`
- `Tests/test_daily_input_transition_evidence.py`

The new `validate_transition_evidence_bundle()` derives obligations exclusively from `required_evidence_for_transition(previous_operation, operation)`. It requires immutable tuples containing exactly one phase report and one transaction receipt per required EvidenceKind; rejects duplicate, missing and extra kinds; validates each report against the same previous/candidate operation; cross-binds report and receipt artifact/verifier fields; checks receipt predecessor hashes in lifecycle-defined evidence order; and verifies the exact artifact bytes under the evidence root.

Also changed `validate_receipt_binding()`: callers must now supply the previous operation, and the function derives the evidence obligations internally. The caller-supplied `required_evidence` parameter has been removed, preventing a caller from selecting a narrower subset.

The bundle return value is explicitly structural, not an authorization token. It does not authenticate verifier origin or a trusted build manifest. Existing verifier registry statuses were not changed: PREPARED remains REPORT_ADAPTER_ONLY and the remaining nine verifier kinds remain NOT_IMPLEMENTED. The complete-bundle synthetic test temporarily substitutes IMPLEMENTED statuses in-memory solely to exercise structural ordering, predecessor chaining and artifact-byte validation; it does not alter the repository registry or establish any real verifier implementation.

### Automated regression result

GitHub Actions run `45597549259`, head `ed3bdec7e3e22a69b1162867f346321503b26498`:
- Windows Python 3.11: **2740 passed, 72 warnings, 44.59s**.
- Ubuntu Python 3.11: **2740 passed, 68 warnings, 28.33s**.
- Both synthetic safety jobs succeeded.
- Overall workflow conclusion: **success**.

Two intermediate CI runs exposed assertion-expectation mismatches in the newly changed tests; the assertions were corrected and the final complete regression above is green. The workflow suppresses warning details; warning counts remain visible and are not claimed to be zero.

This is GitHub Actions synthetic/unit regression only. No local PC test, native Excel COM execution, real XLSM smoke test, trusted verifier attestation, journal append, Apply/Restore enablement or Gate 2 change occurred.

### Updated next steps

1. **Completed:** exact transition evidence bundle validator and internally derived receipt obligations.
2. Define and implement trusted verifier/build-manifest attestation and key custody without inventing or embedding a key.
3. Implement INPUT_AND_LEDGER native Excel COM read-only verifier, reconciled independently with the existing OOXML comparator and the same trusted preflight facts.
4. Continue with the phase-specific verifiers and operational crash-recovery integration only after their trust boundaries are independently satisfied.


## 9. INPUT_AND_LEDGER read-only native COM diagnostic added — 2026-11-25

Implementation `Python/rehab_excel/daily_input_input_ledger_native.py` preserves the independent OOXML comparison and then reads native COM-visible DAILY_INPUT and DAILY_CHANGE_LEDGER from the same saved XLSM. It uses a fresh macros-disabled read-only Excel instance, never saves, and compares the native workbook's full change identities to externally supplied trusted preflight facts. The diagnostic takes fixed read bounds and rejects missing schemas, formula/error cells, wrong visibility/protection, bad identities or changed bytes.

Tests in `Tests/test_daily_input_input_ledger_native.py` use fake COM and synthetic fictional data. This prepares native readback without executing Windows Excel during CI; a disposable native XLSM smoke test is still mandatory when PC is available. INPUT_AND_LEDGER remains NOT_IMPLEMENTED in the phase verifier registry because trusted report/attestation integration and real native proof are absent. This diagnostic is never a journal Apply/Restore authorization.

Ubuntu full CI for code/test HEAD `6e4d99296fc6679fbaeffd4a2003a17ac165719f`: 2770 passed / 77 warnings / 24.10s; Windows CI pending confirmation at this checkpoint. No workbook/VBA/operational writes and Gate 2 stays closed.


### Native INPUT_AND_LEDGER synthetic CI final confirmation — 25/11/2026

GitHub Actions run `45599055363` for code/test HEAD `6e4d99296fc6679fbaeffd4a2003a17ac165719f` finished **SUCCESS** on all four jobs:
- Windows Python 3.11 full regression: **2770 passed**, 77 warnings, 33.58s.
- Ubuntu Python 3.11 full regression: **2770 passed**, 77 warnings, 24.10s.
- Windows and Ubuntu synthetic safety jobs: both success.

Previous interim notes that the Windows outcome was pending are superseded by this final confirmed result. Warning details are suppressed by the CI workflow; do not claim zero warnings. Native COM remains mocked in these tests, not real Excel; no report adapter/trusted attestation/journal append/Apply authorization or Gate 2 activation. An independent disposable-XLSM native Excel smoke test and explicit zero-change policy remain pending.


## 2026-11-25 — INPUT_AND_LEDGER phase report adapter, non-authorizing

Added `Python/rehab_excel/daily_input_input_ledger_phase_adapter.py` and `Tests/test_daily_input_input_ledger_phase_adapter.py`. The adapter requires an exact legal PREPARED -> INPUT_AND_LEDGER_COMMITTED transition, non-empty prior validated preflight facts, a saved XLSM confined within a controlled evidence root, and an exclusive new report path in that root. It executes the existing combined OOXML/native COM read-only input-and-ledger verifier and verifies the source file SHA-256 before and after diagnostic execution. It binds the generated canonical phase report to operation, service date, revision, phase, pinned verifier contract, exact workbook evidence bytes, and native Excel version. Existing report outputs are never overwritten. Invalid phase, inconsistent result, changed workbook, outside-root path and diagnostic failure are rejected before publication.

The INPUT_AND_LEDGER contract status is now **REPORT_ADAPTER_ONLY**, like PREPARED. Eight remaining evidence kinds stay **NOT_IMPLEMENTED**. `validate_phase_verifier_report()` continues to reject INPUT_AND_LEDGER positive authorization while trusted attestation is not integrated; `require_trusted_verifier_attestation()` remains reject-only. Adapter output is neither signed nor a receipt, journal append, actual operational COM proof or Apply/Restore authorization.

Synthetic negative/positive publication tests were added. The adapter tests use an injected fake native diagnostic; they do not execute the user's Excel. GitHub CI results for the current code/test head must be recorded after completion. The prior 2770-pass cross-platform CI belongs to the preceding diagnostic-only commit and must not be misrepresented as coverage of this new adapter. No business workbook, VBA, UI, journal writer or Gate 2 change. Zero-change policy and real disposable native Excel smoke test remain open.


### Confirmed complete regression for INPUT_AND_LEDGER phase adapter — 25/11/2026

GitHub Actions run `45599851668`, code/test HEAD `de03afc8d8cc5d824c45c755d3f0861100a6b278`, completed **SUCCESS**:
- Windows Python 3.11: **2779 passed**, 79 warnings, 39.81 seconds.
- Ubuntu Python 3.11: **2779 passed**, 80 warnings, 17.06 seconds.
- Synthetic safety jobs on both Windows and Ubuntu: success.

This supersedes provisional CI-pending notes. Tests use mock/injected native diagnostic for the phase adapter. No real native Excel/COM execution on the user's machine, actual signed attestation, receipt, positive operational transaction or Gate 2 enablement. INPUT_AND_LEDGER and PREPARED are REPORT_ADAPTER_ONLY. Other eight kinds remain NOT_IMPLEMENTED. Warnings are not zero and warning details are suppressed by CI.

## 10. MASTER projection OOXML structural review — 2026-11-25

The legacy master writer's preview verification is partial, reading A/B/C/status formulas, doctor identity and tail. A new independent saved OOXML comparator `verify_master_projection_saved_readback()` compares the full canonical formula set, two template column schemas, doctor text and row gaps to an earlier expected MasterProjectionRow tuple; hashes source bytes before and after. Fictional synthetic negative tests exercise adversarial changed MASTER cells and expected ordering. This is deliberately **partial PROJECTIONS evidence**, not authorization: the therapist projection and actual native Excel recalculated views remain unverified; PROJECTIONS_REOPEN status unchanged NOT_IMPLEMENTED. Never promote an OOXML-only positive result to trusted native proof.


### MASTER structural readback CI defect and fix — 2026-11-25

The first synthetic regression at code/test commit `0142629049c37dce62c396e55abf0175e6576d88` failed 3 new MASTER positive tests on Ubuntu (2797 passed, 3 failed, 96 warnings). The diagnosis was incorrect literal comparison against Greek text after Unicode `casefold()`: the terminal sigma in `Θάλαμος` and `Ασθενής` was normalized to `σ`, while the comparison constants had `ς`. Fixed both comparisons by casefolding the canonical header constants too in code commit `905418df1c33b4008fd2039a89b6c06d7b4ea7fd`. New Ubuntu Python 3.11 CI `45617783680` passed **2800 tests / 99 warnings / 25.97s**, and both synthetic safety jobs succeeded; full Windows result to be confirmed separately before full green. No real Excel run or workbook modification. This was a newly introduced test/code mismatch, not evidence of operational workbook correction.


### Final cross-platform regression for MASTER structural readback — 2026-11-25

GitHub Actions run `45617783680` for corrected code commit `905418df1c33b4008fd2039a89b6c06d7b4ea7fd`: **SUCCESS** overall, all four jobs successful.
- Ubuntu Python 3.11: **2800 passed, 99 warnings, 25.97s**.
- Windows Python 3.11: **2800 passed, 93 warnings, 42.87s**.
- Both synthetic safety jobs passed.

This supersedes the provisional Windows-pending status for this run. Initial three failing new tests were corrected, not ignored. Warning details are suppressed by CI, counts are retained. This is saved-OOXML verification and synthetic tests only. Full PROJECTIONS_REOPEN remains NOT_IMPLEMENTED and no native Excel, therapist projection, authenticated attestation, Apply or Gate 2 action occurred.

## 11. THERAPIST_DAILY structural view scope — 2026-11-25

Found workload-strip VBA uses provider labels at rows 1 and 11, columns B:J and clock-indexed cells at rows 2–8 and 12–18. Added an explicit saved OOXML diagnostic comparing the exact providers and DISTINCT occupied clock values to independent preflight, without counting session lines as separate occupied slots. This is narrower than the full PROJECTIONS_REOPEN contract. It rejects raw formula occupancy rather than trusting cached values, and requires a later native Excel readback and detailed independently derived session/therapist facts. Missing approved capacity is never converted into zero or six by this verifier. No Gate 2/Apply/Restore changes. CI pending.


### Final cross-platform synthetic CI — THERAPIST_DAILY occupied clocks — 2026-11-25

GitHub Actions run `45618473467`, code/test commit `922afe7bf7e534273781f76ebdf23e600c5ef3bc`, finished **SUCCESS** across all four jobs:
- Ubuntu Python 3.11 full suite: **2815 passed, 98 warnings, 21.23s**.
- Windows Python 3.11 full suite: **2815 passed, 98 warnings, 48.36s**.
- Both Windows and Ubuntu synthetic safety jobs: success.

The warnings are nonzero and CI suppresses their detail. This validates synthetic OOXML readback and static provider/clock expectations only, not native COM, real workbook behavior, full therapist/patient session projection, trusted attestation, phase evidence completion, or Gate 2 permission. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED. No production workbook or VBA changes.


### 2026-11-25 — THERAPIST_DAILY occupied-clock boundary hardening

Review of `daily_input_therapist_projection_readback.py` found that the numeric-time decoder used `round(value * 1440) % 1440`. A valid Excel floating-point time just below 1.0 could round to 1440 minutes and silently become **00:00** instead of failing; this risks false equivalence when checking occupied clocks. Removed the modulo wrap and now reject numeric values rounding outside a 24-hour day. The preflight `expected_clocks` tuple now rejects sub-minute Python `time` values, preventing implicitly unequal resolutions from being compared to minute-based workbook occupancy. Added three synthetic adversarial tests (near-24:00 and two sub-minute expectations) in `Tests/test_daily_input_therapist_projection_readback.py`.

This is a narrowly scoped source/diagnostic hardening, not a clinical capacity-policy change, native Excel proof, authenticated verifier, nor authorization. Source code commit `ee81a24b22f1e139e8a10714f9a09245a0498790`, test commit `d6f3e17046ac208cfd2723fcfac71c6abe37f2d6`; GitHub CI result must be confirmed separately. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED; Gate 2 closed. No workbook or VBA change.


### Confirmed CI: THERAPIST_DAILY clock boundary hardening — 2026-11-25

GitHub Actions run `45619012770`, code/test head `d6f3e17046ac208cfd2723fcfac71c6abe37f2d6`, completed **SUCCESS** overall:
- Windows Python 3.11: **2818 passed, 92 warnings, 29.70s**.
- Ubuntu Python 3.11: **2818 passed, 96 warnings, 27.97s**.
- Synthetic safety jobs on both platforms: SUCCESS.

The added clock-wrap/sub-minute tests passed. This is synthetic saved-OOXML verification only; no real Excel COM/native recalculation, authenticated attestation, complete provider-session proof or production changes. Warning counts remain nonzero and details are suppressed by CI. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED; Gate 2 closed.


### 2026-11-25 — THERAPIST_DAILY assignment identity and shared-slot membership contract

Added `Python/rehab_excel/daily_input_therapist_assignment_membership.py` and `Tests/test_daily_input_therapist_assignment_membership.py`. The **pure** `verify_therapist_assignment_memberships()` comparator requires two separately sourced immutable tuples of `TherapistAssignmentFact` (preflight expected and saved-readback observed) keyed by *stable* immutable AssignmentID, each binding PatientID, treatment, normalized provider, explicit service date and exact-minute start time. It rejects missing, duplicate, swapped, stale, legacy/row-derived, cross-date or sub-minute facts even if occupied-provider-clock counts are unchanged. For same-provider simultaneous different patients, the exact group of patient IDs must match independent prior approved shared-slot membership; excess/missing/duplicate approvals fail closed. It does not invent provider limits or automatically authorize shared treatments.

**Important trust boundary:** the comparator cannot authenticate its caller or itself reconstruct actual patient assignment IDs from the display cells of THERAPIST_DAILY. Its observed tuple must be produced by a separate independently trustworthy saved-source/COM verifier in later work. The comparator is **not** a completed positive projection verifier, report adapter, trusted attestation or Apply/Restore authorization. No workbook IO, VBA mutation, journal append or Gate 2 activation. `PROJECTIONS_REOPEN` stays NOT_IMPLEMENTED. This pure unit test suite tests in-memory fictional identity sets; it cannot prove real native Excel or full clinical assignment correctness.

GitHub code/test commits `f15e366ea01c436d46c939c05709f16cc6b33cd0` / `166d7bf54345476422dad79e2c4de964891c9214`. CI outcome must be separately recorded once confirmed.


### Confirmed cross-platform assignment membership regression — 2026-11-25

GitHub Actions run `45622874406`, code/test head `166d7bf54345476422dad79e2c4de964891c9214`, finished **SUCCESS** in all four jobs:
- Windows Python 3.11: **2839 passed, 93 warnings, 60.67s**.
- Ubuntu Python 3.11: **2839 passed, 93 warnings, 19.48s**.
- Windows and Ubuntu synthetic Python safety jobs: success.

Added tests reject changed immutable AssignmentID, patient identity, treatment, provider, date, time, duplicate assignment, missing row, legacy ID, sub-minute drift, missing/extra/shared-approval mismatch. Tests are fictional in-memory examples, **not** a real workbook source or native Excel COM inspection. No patient-identifying data was copied into repository fixtures. The complete PROJECTIONS_REOPEN phase, including independent saved-source extraction/attested report, remains NOT_IMPLEMENTED and Gate 2 remains closed. CI warnings remain nonzero.

 
## 12. Partial native MASTER identity readback added — 25/11/2026

Added `Python/rehab_excel/daily_input_master_projection_native.py` and `Tests/test_daily_input_master_projection_native.py`. This new diagnostic first invokes the existing strict saved-OOXML MASTER structural comparator, then opens the exact saved XLSM in a fresh isolated Excel COM instance with macro security forced disabled, events/alerts/link updates disabled and ReadOnly=True. It checks native formula linkage and Excel-evaluated room, patient display identity and status against independently supplied immutable `MasterProjectionRow` facts; doctor-bearing identity is checked as the established multiline literal. It captures Excel Version, closes with SaveChanges=False, quits in finally and rejects cleanup failure or changed source bytes.

CI run `45657152050`, HEAD `e7ce7526de2c000cff9410819267f73fdc6ffc79`: Windows **2848 passed / 35 warnings / 48.64s**; Ubuntu **2848 passed / 34 warnings / 28.83s**; both synthetic safety jobs succeeded.

The fake-COM tests do not execute Microsoft Excel. This is not a full projection verifier: it does not independently reconstruct native therapist assignment IDs, prove all therapist-day projections or source-ledger linkage, verify every evaluated MASTER treatment value/style, produce a phase report/receipt or authenticate verifier provenance. `PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`. The verifier registry was not changed. No workbook/VBA/operational handler, journal writer, Apply/Restore or Gate 2 change.

Detailed boundary and ordered next steps: `Documentation/DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md`.

 
## 13. Therapist assignment source-chain audit — 25/11/2026

A follow-up source-only review mapped existing code that should be reused for the next independent THERAPIST_DAILY reader:
- stable ID binding: `stable_recurring_ids.bind_stable_assignment_ids()`;
- canonical inpatient/granular source: `reader.read_base_schedule()` with `PROGRAM_WEEKLY_SLOTS` overlay;
- unified inpatient/outpatient source: `outpatient_schedule_source.read_unified_base_schedule()`;
- dated session identity: `materialize_sessions_for_date()` builds `{base_entry_id}@{service_date}`;
- daily replacement source: `DAILY_REPLACEMENT_FACTS` records exact TargetSessionId, original/effective provider, date/time, provider kind and status;
- effective daily state: `build_daily_session_states()` preserves original session identity and computes the effective provider/time/status.

The visible THERAPIST_DAILY grid contains only provider/time/patient occupancy display, not AssignmentID. Therefore the future native verifier must derive expected effective assignments from the authoritative native-captured sources and resolve each displayed patient/provider/time to exactly one assignment. Missing or ambiguous joins must fail closed. It must preserve exact weekly/outpatient source row indices for stable-ID binding, include PATIENT_EXITS/recurring changes/daily input ledger/replacement overlays, and separate student from therapist replacements. It must not infer IDs from labels, ordering or occupied-clock counts.

This is a source map and design constraint, not an implementation. Full detail and unresolved label/zero-day questions: `DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md` §7. No verifier registry status changed; PROJECTIONS_REOPEN remains NOT_IMPLEMENTED.

 
## 14. Native THERAPIST_DAILY full-grid readback primitive — 25/11/2026

Added `daily_input_therapist_projection_native.py` and `test_daily_input_therapist_projection_native.py`. The diagnostic requires exactly 18 provider headers and explicit expected text/blank values for all 126 patient cells in B2:J8 and B12:J18. It opens a fresh macro-disabled Excel COM instance read-only, checks provider headers and the fixed time grid, compares all native cell values exactly, rejects formulas/errors/non-text values, captures Excel Version, closes without saving, quits and verifies unchanged workbook bytes. It deliberately does not infer AssignmentIDs from display labels.

GitHub Actions run `45658828780`, HEAD `71ff839699782615e9968d4213c9b8d42ea90bc6`: Windows **2859 passed / 32 warnings / 44.28s**; Ubuntu **2859 passed / 35 warnings / 19.52s**; both synthetic safety jobs passed.

This is a fake-COM tested primitive, not an actual Excel run and not a complete PROJECTIONS_REOPEN verifier. The full 126-cell expected-grid generator, independent AssignmentID observed-membership extraction, shared-slot proof and operation-bound DAILY_INPUT/ledger linkage remain unimplemented. No phase report, receipt, trusted attestation or authorization. Registry unchanged; Gate 2 closed. Details: `DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md` §8.

 
## 15. Full expected THERAPIST_DAILY grid builder — 25/11/2026

Added `daily_input_therapist_projection_expected.py` and its tests. The builder requires immutable stable-ID-bound base entries, dated sessions, exact one-to-one DailySessionState coverage, patient/student registries and independently supplied provider headers. It reuses `render_base_slot()` and `build_patient_centric_preview_plan()` to create the complete expected 126-cell display, including existing absence/replacement annotations. It separately derives expected active therapist AssignmentID facts, excluding absent/cancelled and student-replaced sessions. It rejects duplicate/unstable AssignmentIDs, state/session mismatch, unknown patients, ambiguous duplicate display names in a visible cell and StudentID/TherapistID collisions.

A synthetic student replacement exposed that the existing preview's final provider/time patch pass had discarded provider kind. Corrected `patient_centric_preview._provider_cell()` to recognize a known StudentID during final resolution. Replaced stale `student.student_number` references with `student.student_id` in `patient_centric_preview.py` and `outpatient_presentation.py`.

GitHub Actions run `45660810319`, HEAD `5a91058f0c9f78f63c19a70366987d8685796af7`: Windows **2866 passed / 36 warnings / 42.21s**; Ubuntu **2866 passed / 32 warnings / 23.47s**; both synthetic safety jobs succeeded.

The builder consumes caller-supplied facts. It does not independently capture or validate the source sheets, stable-ID registry, exits/recurring changes, DAILY_INPUT, authoritative ledger or replacement store. It is not joined to the native full-grid COM reader in a composite operation-bound diagnostic/report adapter. No trusted attestation or registry status change; PROJECTIONS_REOPEN remains NOT_IMPLEMENTED; Gate 2 remains closed.
