# DIOMIL — NEXT SESSION HANDOVER

**Repository:** diomil73/EKA-REHAB-TIMETABLE  
**Branch:** work/afternoon-medical-scope-20261115  
**Last operator-confirmed local HEAD:** b2be9373fe1bc3695eaf6f0eb075e5f763c2e1e6  
**Last checked GitHub branch HEAD:** b2be9373fe1bc3695eaf6f0eb075e5f763c2e1e6  
**Date:** 2026-11-28

## Read first, in this order

1. Documentation/DIOMIL_MASTER_CONTINUITY_GUIDE.md — authoritative full continuity and operator instructions; continue from latest §4.253.
2. Documentation/POWERSHELL_OPERATOR_EXECUTION_STANDARD.md — mandatory safe execution method; read before giving any Windows command.
3. This file — exact current location and next action.
4. Documentation/CURRENT_HANDOFF.md — detailed technical state of daily-input verifiers and projection work.
5. Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md — delivery history and evidence boundaries.
6. Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md — exact test/CI evidence and non-authorizing boundaries.
7. For projection details: Documentation/DAILY_INPUT_PROJECTIONS_NATIVE_READBACK_DESIGN_20261125.md.
8. For the eventual real-data DEMO: Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md and Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md.

## What is already GREEN locally

- Safe Git fast-forward completed; local and origin branch were equal at b2be9373. Tracked working tree was clean. Untracked logs/backups/previews were preserved.
- PREPARED_SNAPSHOT focused test: 1 passed, exit 0.
- Targeted Python 3.14 warning diagnostic across five files: 109 passed, 1 skipped (Windows symlinks unavailable), exit 0. No ZipFile warnings appeared in the supplied output.

Do not rerun these without a specific technical reason.

## Current unresolved PC task

Identify which of the three frozen files, if any, has the required expanded MASTER K/L/M and PATIENT_PLANNER psychology headers, using read-only OOXML only:

- Excel/previews/UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm
- Excel/previews/UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm
- Excel/previews/UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm

The earlier schema comparison attempt on 2026-11-28 did not complete. The operator transcript shows split PowerShell execution: finally is not recognized, transcript remained open, and moving the active transcript failed. The printed FINAL_STATUS=B is invalid. None of the three candidates has been classified BAD by that attempt.

The earlier source Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm is not suitable for the new K preview: the actual preflight correctly stopped because PATIENT_PLANNER lacks Ψυχ_Ώρα. Do not rerun its preview tool.

## Next action

Perform only the read-only schema comparison of the three frozen candidates, with baseline hash checks and exact header positions. Use the saved-script or one-paste block method in POWERSHELL_OPERATOR_EXECUTION_STANDARD.md. Do not repeat the structural auditor unless a concrete issue requires it. Do not open or modify any frozen workbook.

Only after a candidate is identified should its provenance/role be checked against the existing guide and matrix before creating a new uniquely named disposable native preview.

## Naming and status rules

New evidence filenames use descriptive purpose + timestamp + status suffix:
- _G = diagnostic completed and GOOD.
- _B = diagnostic completed and BAD.
- _P = pending/unreviewed.
- _S = wrapper failure or safety stop.

A wrapper failure is never a candidate _B.

## Guardrails

- No production workbook changes.
- No frozen candidate overwrite, repair, deletion or retry.
- No Gate 2 opening or positive operational transaction.
- Separate Python GREEN, native Excel GREEN and operational acceptance.
- Do not repeat already-green local tests without technical cause.
- Update this handover, master guide, delivery report, matrix and current handoff after every meaningful checkpoint.


## Latest engineering checkpoint — 2026-11-28

The mobile-safe DAILY_INPUT evidence work has advanced beyond the previous PROJECTIONS_REOPEN report adapter. Added:
- \`Python/rehab_excel/daily_input_pending_decision_inventory.py\`
- \`Tests/test_daily_input_pending_decision_inventory.py\`

It creates a deterministic, operation-bound snapshot of every unresolved therapist-absence target from the existing projection preflight state model, validates exact session/state identity, counts all five statuses, distinguishes original operation source hash from current evidence-workbook hash, and writes a new exclusive evidence artifact. It does not include provider options or decisions and is not a native verifier/report adapter.

Final synthetic full CI: run **46161976361**, exact code/test HEAD \`bf8c6e34c71c45208f8312ff857aafe698cff07c\`: Windows 2958 passed / 28 warnings; Ubuntu 2958 passed / 27 warnings; both safety jobs passed.

**Next mobile engineering task:** implement the separate PENDING_DECISION_INVENTORY native-readback phase-report adapter using the existing composite native source-capture diagnostic. Keep it REPORT_ADAPTER_ONLY and non-authorizing. Then continue to REPLACEMENTS_REOPEN. Do not repeat the already-green full regression absent another code change.

**PC task remains separate:** read-only schema comparison of the three frozen H10 candidates; previous split PowerShell wrapper result is invalid and no candidate was classified BAD. No frozen workbook changed.


## Latest engineering checkpoint — 2026-11-28

The PENDING_DECISION_INVENTORY phase-report adapter is now implemented:
- \`Python/rehab_excel/daily_input_pending_decision_phase_adapter.py\`
- \`Tests/test_daily_input_pending_decision_phase_adapter.py\`
- Strict canonical inventory decode/readback functions in \`daily_input_pending_decision_inventory.py\`.
- Registry status for this evidence kind is **REPORT_ADAPTER_ONLY**.

The adapter reopens the exact inventory artifact, checks canonical bytes/digest, rebuilds the expected inventory from preflight and compares it, binds operation/source/date/workbook digest, reuses the existing composite native read-only Excel verifier, checks source/projection counts and native Excel version, then publishes and reopens a canonical phase report exclusively. It does not authorize the phase, write workbook/VBA, append a journal or receipt, or enable Apply.

Final synthetic CI run **46162899470**, exact code/test HEAD \`1b90e3dab9e5dc3d4bd9166e5c321ab665495cde\`: Windows **2964 passed / 30 warnings**, Ubuntu **2964 passed / 28 warnings**; both synthetic safety jobs passed. An earlier run found one outdated registry-test expectation; it was corrected before final GREEN. CI native readback is injected/fake, not actual Excel COM.

**Next mobile engineering task:** implement REPLACEMENTS_REOPEN report adapter using the existing replacement facts and suspension readers, with exact operation/date/workbook binding and synthetic negative tests. Do not implement replacement choices/operator decisions yet.

**PC-only task:** read-only schema comparison of the three frozen H10 candidates remains pending. Previous split PowerShell wrapper's FINAL_STATUS=B is invalid. No frozen workbook was modified. Local PC HEAD after latest GitHub commits is not yet operator-confirmed.


## Latest engineering checkpoint — 2026-11-28

REPLACEMENTS_REOPEN report adapter is implemented:
- \`Python/rehab_excel/daily_input_replacements_reopen.py\`
- \`Python/rehab_excel/daily_input_replacements_reopen_phase_adapter.py\`
- \`Tests/test_daily_input_replacements_reopen_phase_adapter.py\`
- Contract registry status: **REPORT_ADAPTER_ONLY**.

It reopens the prior pending inventory and verifies that every pending target resolves exactly once to an active replacement fact or explicit treatment-suspension cancellation. It checks the immutable session/patient/original therapist/date join, replacement provider/time against final state, suspension ID/kind/date against final cancellation state, and reuses native source capture for DAILY_INPUT and DAILY_REPLACEMENT_FACTS. Evidence and phase report are canonical, exclusive and read back.

Final synthetic CI run **46165623512**, exact code/test HEAD \`ad18f3305bd3d5f3a3caa7c1f6b55df575940f5a\`: Windows **2972 passed / 29 warnings**, Ubuntu **2972 passed / 28 warnings**; both safety jobs passed. CI injects fake native capture; no actual Excel COM run is claimed. OPERATOR_RESOLUTION remains a separate proof obligation; this report is not trusted attestation or authorization.

**Next mobile engineering task:** implement OPERATOR_RESOLUTION evidence artifact/reader bound to the exact operation ID, pending inventory digest, resolution mode and target set. Then integrate OPERATOR_RESOLUTION and REPLACEMENTS_REOPEN together in the phase evidence flow. Do not infer operator approval from the lifecycle enum alone.

**PC-only task:** read-only schema comparison of the three frozen H10 candidates remains pending. Prior split PowerShell wrapper's FINAL_STATUS=B is invalid. No frozen workbook was modified. Local PC HEAD after latest GitHub commits is not yet operator-confirmed.

## 2026-11-28 — H10 K schema comparison diagnostic added; local frozen-file read pending

The previous schema comparison wrapper failure remains a **wrapper SAFETY STOP**, not a candidate failure. No candidate was classified BAD and the frozen workbooks remain untouched.

Added on the active branch:
- `Python/rehab_excel/h10_k_schema_compare.py`: pinned baseline manifest and read-only OOXML schema comparison helpers.
- `Python/tools/diagnose_h10_k_schema_compare.py`: one-command diagnostic for the three frozen candidates.
- `Tests/test_h10_k_schema_compare.py`: synthetic tests for the pinned manifest and MASTER header-row detection.

The tool uses the existing `WorkbookPackageProbe`, checks each candidate against the exact SHA-256 and byte-size recorded in `INTEGRATED_TEST_BASELINE_MATRIX_20261115.md`, verifies ZIP integrity and VBA presence, then checks:
- MASTER header row identified by the exact B/C `Θάλαμος` / `Ασθενής` pair, with exact K/L/M labels.
- Exactly one row-1 occurrence of each PATIENT_PLANNER header: `PatientID`, `Ψυχ_Ώρα`, `Ψυχ_Ημέρες`, `Ψυχ_Ψυχολόγος`.
- Source SHA-256 and size unchanged after inspection.

It prints no patient values. It does not launch Excel/COM, save or modify a workbook, select an integration baseline, prove synthetic provenance, or grant native/operational acceptance. `_G` means exactly one schema match; `_B` means the completed comparison found zero or multiple matches; `_S` means a safety/baseline/read failure.

Three GitHub commits currently contain the files:
- `963f9968bdd9c96fa3f558fe69a64e607b1a2420`
- `f625ecac3b0ad6fe91373e2306e445f4875f608e`
- `54c104fb138a526deb20f0f1ce4adee3faeac956`

At this documentation update, GitHub Actions synthetic safety runs for those commits were still in progress. No local Windows test or frozen-file comparison has yet been executed. Do not call this Python GREEN, native Excel GREEN, or candidate selection.

**Next PC action:** fast-forward the local branch, run only `Tests/test_h10_k_schema_compare.py`, then run `Python/tools/diagnose_h10_k_schema_compare.py`. Preserve the generated report and return its complete output. Do not rerun the older structural auditor, the previously green receipt tests, or any frozen workbook operation.


## 2026-11-28 — H10 K local SAFETY_STOP root cause: VBA member case mismatch

Operator local HEAD `dd8d9daf337784e9d22d5d71f8a0752b859356ce` after fast-forward; focused `Tests/test_h10_k_schema_compare.py`: **3 passed in 2.56s**. Read-only diagnostic report `DIOMIL_H10_K_SCHEMA_COMPARE_20261128_184701_S.txt` returned `FINAL_STATUS=S`, exit 2, `CANDIDATES_COMPLETED=0/3`, with `VBA_PROJECT_MISSING` for all three frozen candidates. **This is NOT a candidate BAD classification and NOT evidence that VBA is absent.** No frozen workbook modified.

Code review found a deterministic diagnostic bug: ZIP entry names were `casefold()`-normalized but compared against mixed-case `xl/vbaProject.bin`; thus even a valid `xl/vbaProject.bin` falsely returned missing. Fixed case-consistent lookup via `has_vba_project_member()` in `Python/rehab_excel/h10_k_schema_compare.py` and added focused positive/negative regression tests in `Tests/test_h10_k_schema_compare.py`. Code commit `18e837fe8f8365bea8d066962a7f908b7196d92b`, test commit `1d49ee83cdc30c31a819274f5cf85b2b4b5e2432`. **New code not yet locally tested on operator PC; CI outcome must be verified separately.** This corrects only the diagnostic; it does not affect production workbooks, permissions or Gate 2.

**Next PC action:** fast-forward branch, run only `python -m pytest -q Tests/test_h10_k_schema_compare.py`, then `python Python/tools/diagnose_h10_k_schema_compare.py`. Preserve full result and report. If status is `S`, stop and investigate; if `G` or `B`, evaluate schema findings only, not provenance or operational acceptance. Do not repeat already GREEN unrelated tests.


## 2026-11-28 — Operator-confirmed H10 frozen K schema comparison: completed, zero matches

Operator fast-forwarded local branch `dd8d9daf` -> `b4d85e2e` and ran only the focused H10 schema tests: **5 passed in 1.41s**. Completed the read-only OOXML diagnostic `DIOMIL_H10_K_SCHEMA_COMPARE_20261128_185237_B.txt` with `DIAGNOSTIC_EXIT_CODE=0`, `CANDIDATES_COMPLETED=3/3`, `SCHEMA_MATCH_COUNT=0`, `UNIQUE_SCHEMA_CANDIDATE=NONE`, `FINAL_STATUS=B` (SCHEMA_ONLY, not provenance or acceptance).

All three pinned frozen H10 candidates matched their exact expected SHA-256 and size, passed ZIP integrity and VBA presence, had MASTER row 4 exact B/C/K/L/M headers, and had unique PATIENT_PLANNER PatientID in column A. **All three lacked exact row-1 PATIENT_PLANNER headers `Ψυχ_Ώρα`, `Ψυχ_Ημέρες`, `Ψυχ_Ψυχολόγος`.** Therefore MASTER schema matched and PATIENT_PLANNER expanded psychology source schema did not; no candidate can be designated as the unique expanded-K integration baseline. No frozen workbook was modified; no patient values were printed, no Excel COM was launched. This does not invalidate historical positive native tests for their previously scoped workflows.

**Next engineering step:** do not repeat this frozen candidate comparison. Trace the existing approved PATIENT_PLANNER field/source and workbook-construction pathway, then propose a separate disposable schema-compatible derivative with exact new headers and preservation of legacy contents, macros and provenance controls. Never alter frozen sources or claim native acceptance from synthetic tests. If new local Excel work is required, isolate that PC step. Gate 2, receipts and Apply/Restore remain closed. Mobile-safe OPERATOR_RESOLUTION evidence remains separately pending.


## 2026-11-28 — PC psychology preview path: reuse existing builder, disposable only

After operator-confirmed frozen H10 schema result B (all 3 lack exact PATIENT_PLANNER psychology fields), the existing `Python/tools/preview_psychology_schema.py` was inspected. It ALREADY appends `Ψυχ_Ώρα`, `Ψυχ_Ημέρες`, `Ψυχ_Ψυχολόγος`, adds `SETTINGS!PSYCHOLOGISTS`, clones the existing extended x14 Excel validation nodes, and verifies the dropdowns; it copies source to a distinct preview first. Do not create a redundant builder. Security hardening commit `d28d0aa634a0c26cdbd4002afb3ae5f9f634d27b` sets `Excel.Application.AutomationSecurity=3` before both writable-copy and read-only-validation COM opens. This is not yet local/native verified.

**Next operator PC action (not yet performed):** run a single guarded PowerShell block that fast-forwards the branch; checks the pinned SHA-256 of the `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` candidate; refuses any existing output; runs the existing preview tool **without `--overwrite`** on the candidate as an explicitly provisional, disposable illustration only, not a selected integration baseline; then prints status. Keep frozen originals untouched. Do not label any preview as business accepted until actual readback and visual/manual acceptance. Do not infer that the preview has valid clinical provenance or real-data approval. Preserve generated output/report. Gate 2 CLOSED.


## 2026-11-28 — Psychology disposable preview native success; VBA comparison pending

Operator confirmed `PSYCHOLOGY SCHEMA PREVIEW OK` after fast-forward to `930dbf0ce56c042b9a6b21884e9f58a243524d9b`. Distinct, disposable `Excel/previews/H10_CC_PSYCH_SCHEMA_DISPOSABLE_20261128.xlsm` was built from pinned frozen CC source. PATIENT_PLANNER appended `Ψυχ_Ώρα` column 24, `Ψυχ_Ημέρες` column 25, `Ψυχ_Ψυχολόγος` column 26; SETTINGS `PSYCHOLOGISTS` column 9; three native Excel dropdowns verified. `PREVIEW_EXIT_CODE=0`, `FROZEN_SOURCE_UNCHANGED=YES`, VBA present, but `VBA binary identical=False`. This difference alone does not prove a semantic VBA regression nor equivalence; the preview is NOT operationally accepted and is not a replacement for the frozen source.

New read-only diagnostic `Python/tools/diagnose_psychology_preview_vba_semantics.py` (commit `e6b9cc4a9fed028975726cc44e487c84de512dc4`) reuses `audit_vba_components` and `assert_vba_semantics_preserved`, opening original and preview under macro-disabled native COM with ReadOnly=True and SaveChanges=False, checking source pinned digest, component code/export/resources and byte invariance. It is not yet executed on operator PC. It may require the Excel Trust Center option allowing programmatic VBA project access; do not automatically change security settings. Fail-closed result is not proof that preview is corrupt. Next PC work: run this **single** focused diagnostic without rerunning the completed H10 schema comparison or preview builder. Preserve both XLSM unchanged. Gate 2 CLOSED.


## 2026-11-28 — Psychology preview VBA resources: 10 exported-resource differences, next independent readback

Operator PC ran the focused read-only diagnostic at local HEAD `66424f85ed67d111fb7e315e0122e323b02e6599`: `VBA_COMPONENT_SEMANTICS=SAFETY_STOP`, `Preexisting VBA component mismatch: missing=0 identity=0 code=0 export=0 resource=10`, `DIAGNOSTIC_EXIT_CODE=2`. Frozen source and disposable preview both remained byte-identical to their pre-inspection SHA-256 (`UNCHANGED=True`); macros not executed, workbook not saved. This proves same module identity, code and exported component files, **not** UserForm resource preservation. Ten exported FRX companions differ; unchanged VBA code does not justify a GREEN classification.

Existing independent embedded OLE stream and read-only UserForm designer auditors were reused in commit `b9c302968d4906322586c2736598553005b33f5f`, enhancing only `Python/tools/diagnose_psychology_preview_vba_semantics.py` to print privacy-safe aggregate stream-change categories and check designer parity. Strict VBA resource integrity remains enforced: even if the designer parity is GREEN, the diagnostic does NOT silently waive FRX differences or authorize clinical use. If inspection fails, retain fail-closed and investigate. No need to rebuild the preview or repeat earlier tests. Next PC action: `git pull --ff-only` then run the same single diagnostic; capture complete output. Gate 2 CLOSED.


## 2026-11-28 — Psychology preview: embedded OLE and designer parity established; final verifier wired

Operator native read-only comparison at HEAD `3738733d2e19549a8ed14fc9ab2d5e233693c036` reported **10 UserForms**, OLE forms/streams added, missing or changed: **all zero**, and `USERFORM_DESIGNERS_MATCHED=10`. Frozen source and disposable preview remained unchanged; no macros or saves. The old `VBA_COMPONENT_SEMANTICS=SAFETY_STOP` was solely the 10 differing Excel-exported `.frx` companion digests. This is not evidence of changed embedded UserForm storage or code.

Existing `assert_vba_semantics_preserved` already supports `independently_verified_form_storages=True` but only AFTER independent exact embedded OLE proof. Code commit `3851fc76b69d3486cfd2467709143a6bcd06b0ab` now calls strict `assert_form_storage_preserved(source_ole, preview_ole)` (all UserForm stream sets and bytes), requires designer parity and passes that proof flag to the existing VBA component checker. It continues rejecting missing/added components, changed code, changed exported VBA modules and unexpected non-UserForm resource changes. This is a scoped verifier wiring fix, not relaxation without evidence. Existing synthetic tests `test_form_frx_export_variation_needs_independent_raw_ole_proof`, `test_independent_form_proof_does_not_ignore_module_export_or_code_changes` and `test_form_storage_refuses_real_stream_changes` already cover the trust rule; do not rerun full suites without reason.

**Next PC step:** fast-forward and execute only `Python/tools/diagnose_psychology_preview_vba_semantics.py` to verify the revised wiring against the unchanged local pair. No preview rebuild, no frozen edits, no Gate 2. Even GREEN would mean only VBA semantic preservation, not whole-workbook functional/clinical acceptance.


## 2026-11-28 — PC psychology disposable preview VBA semantics verified GREEN

Operator completed native read-only VBA semantic verification at local HEAD `bb98a08689320846ed316b1532811f227ed587a1`, exit code **0**. Evidence: `USERFORMS_AUDITED=10`; OLE forms added/missing, forms with changed streams, streams added/changed/missing, f/o/other stream differences all **0**; `USERFORM_OLE_STREAMS_PRESERVED=10`; `USERFORM_DESIGNERS_MATCHED=10`; `VBA_COMPONENT_SEMANTICS=GREEN`; `COMPONENTS_MATCHED=66`; `FROZEN_SOURCE_UNCHANGED=True`; `DISPOSABLE_PREVIEW_UNCHANGED=True`; `NO_MACROS_NO_SAVE=YES`. This closes the preview VBA semantic and embedded UserForm preservation verification. The previous 10 differing exported FRX companions were permissible under the existing explicitly independently verified exact OLE-stream preservation clause. Do NOT repeat the same check absent material change.

The local disposable `Excel/previews/H10_CC_PSYCH_SCHEMA_DISPOSABLE_20261128.xlsm` still requires safe manual visual/functional review. The preview contains the appended PATIENT_PLANNER psychology headers in X:Y:Z and SETTINGS PSYCHOLOGISTS in I and native dropdown validations, but **no real clinician roster acceptance, clinical approval, full K content or complete workflow integration is claimed**. Next PC task: open ONLY the disposable preview for visual read-only review, with VBA/macros disabled, no save, no patient-value transcription or edits; inspect presentation of MASTER K and psychology columns/dropdowns. Record operator observations, then select precise next integration changes. Frozen sources remain immutable; Gate 2 CLOSED.


## 2026-11-28 — Operator PC visual screenshots of psychology disposable preview

Operator supplied three actual Windows Excel screenshots of `H10_CC_PSYCH_SCHEMA_DISPOSABLE_20261128.xlsm`: (1) MASTER_SCHEDULE overview, (2) left PATIENT_PLANNER treatment fields, (3) right PATIENT_PLANNER columns including psychology. Observations only, not proof of functional data propagation: main MASTER layout, coloured treatment groups and top navigation buttons are visible; MASTER K `ΨΥΧΟΛΟΓΟΙ` and L `ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ` visible and blank in displayed rows; PATIENT_PLANNER trailing psychology headers `Ψυχ_Ώρα`, `Ψυχ_Ημέρες`, `Ψυχ_Ψυχολόγος` visible. The new header for psychologist is visually blue while preceding psychology header cells are peach, with narrow/cropped adjacent older headers. Existing visible patient data should not be reproduced into public logs. Do not assert that blank K is a code defect without a synthetic populated source. PM L remains deliberately deferred pending approved layout. Previous native VBA semantic/OLE evidence GREEN stands and should not be rerun absent changed path.

Next: use the existing psychology VBA renderer and MASTER K bulk projection contract to design one disposable fictional populated source-to-MASTER K test, and a separate minimal visual-format adjustment when warranted; no real patient edit, frozen source modification, automatic approval, Gate 2 or operational transaction. The displayed workbook is not final clinical acceptance.
