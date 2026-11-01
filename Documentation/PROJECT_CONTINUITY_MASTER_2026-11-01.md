# Project Continuity Master — 2026-11-01

This document is the authoritative continuity anchor for the EKA Rehab Timetable project.

It exists so that a new ChatGPT conversation, developer, or long-gap return can resume without losing the accumulated technical knowledge, without repeating failed experiments, and without drifting away from the operational goals of the rehabilitation center.

If context is lost, read this document before making changes.

---

## 1. Project mission and non-negotiable goals

This is not merely an Excel automation project. The system is intended to become the operational digital backbone of the rehabilitation center.

The primary goals are:

1. Simplify and solve, through digitization, the operational handling of patients when they:
   - enter the program,
   - leave/discharge,
   - are absent,
   - return,
   - change treatments,
   - change therapists,
   - need temporary replacement/reassignment,
   - experience cancellations or treatment suspension.

2. Provide a complete and trustworthy picture of what is happening in the center every day.

3. Preserve complete history and statistics for both patients and therapists.

4. Track patient progress over time, including milestone/progress evaluation.

5. Provide easy linkage between:
   - the operational timetable,
   - the patient record/folder,
   - the patient progress/history view.

These goals are the product compass. A technical change that improves a local mechanism but weakens these goals is not progress.

---

## 2. Operating philosophy

The application should behave like a coherent operational system, not a collection of disconnected Excel sheets.

The workbook is the user-facing application surface. Python contains much of the business/scheduling logic. VBA provides the in-workbook interaction and transaction bridge.

The system must preserve one daily operational truth and must not create competing truths in different sheets.

Locked daily truth:

`recurring/base schedule + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> operational views/statistics/history`

The daily recomputation order is conceptually:

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. resolve remaining orphans through valid placement or treatment suspension,
7. only then expose remaining capacity to new patients.

A finalized day must have unresolved/orphan sessions = 0.

---

## 3. Safety rules that must never be forgotten

These rules are mandatory.

- The production/baseline workbook `Rehab_Center_System_v27_1.xlsm` is immutable.
- Never experiment directly on the production workbook.
- Never reuse a failed native smoke workbook for another repair attempt.
- Every native smoke scenario uses a fresh uniquely named `.xlsm` copy.
- Failed artifacts are frozen as evidence.
- Do not do repair-over-repair on the same workbook.
- Preserve VBA.
- Verify the source workbook SHA256 before/after build operations.
- Do not let openpyxl save production/staging workbooks in a way that may damage VBA.
- Source changes need Python regression tests before native Excel proof.
- Native Excel proof is required before considering an operational behavior green.
- Diagnostic-first. No blind alphabetic rebuild loops.
- Do not modify the replacement/global solver unless evidence points there.
- Do not discard proven business logic merely because a newer transaction layer fails.
- Windows tests use `python -m pytest`, not bare `pytest`.

Known-good hard recovery boundary:

`4a2eed09ae0073ea61d34d24d55e5e340269d3a6`

Immutable tag/branch references were created around this boundary, including:

- `baseline/known-good-4a2eed0`
- `backup/pre-handoff-recovery-2026-10-29`

Important: this boundary is a safety/recovery baseline, but it is NOT the last point where the later replacement/patient-centric engine existed. Do not confuse those two concepts.

---

## 4. Current working branch and repository state

Active branch:

`work/history-statistics-rebuild`

The branch contains the recovery work, day-history work, close-day architecture, in-place DAILY_INPUT transaction work, operator visibility, restore ledger, and the current diagnostics.

The branch is substantially ahead of the hard recovery boundary.

The older `Documentation/CURRENT_HANDOFF.md` was stale as of 2026-11-01 and must now point to this document.

---

## 5. User workflow preferences that matter for continuity

The user wants the assistant to drive progress proactively.

Standing collaboration rules:

- Continue without waiting for repeated “ok”.
- Do not ask redundant questions.
- When a local/native result is genuinely required, give exact PowerShell commands and expected results.
- Otherwise continue the investigation autonomously.
- Update repo notes/checkpoints for every meaningful source change, native outcome, failure mode, or architectural decision.
- Do not restart old debugging tracks without evidence.
- Preserve the project history and use checkpoints as evidence.
- Prefer partial concrete progress over asking for clarification when the evidence already points to a safe next step.

---

## 6. Product roadmap and target architecture

The broad roadmap remains:

1. Scheduling engine
2. Day closure/history truth
3. Statistics/reporting
4. Patient progress
5. App-like UI
6. Integrated beta
7. Hardening
8. Production

Expanded functional roadmap includes:

- patient lifecycle,
- therapist history/statistics,
- department statistics,
- patient history,
- 15-day progress milestones,
- aggregated reporting,
- doctor/clinic identity where needed,
- holiday/attendance semantics,
- reporting UI,
- MASTER as home/dashboard,
- patient-folder links,
- final UI polish/hide Excel chrome,
- optional later waitlist,
- future clinical notes gate.

The current blocker is before day closure/history can safely advance: DAILY_INPUT apply/restore must first be natively reliable in-session.

---

## 7. History/statistics decisions already locked

Historical truth must not be reconstructed later from today’s mutable workbook projection.

Locked persistence model:

- finalized daily snapshots for fast reporting,
- append-only historical events for audit/reconstruction.

Closed days use explicit `Κλείσιμο ημέρας`.

Corrections use an auditable amendment/versioning flow such as `Διόρθωση κλεισμένης ημέρας`, not destructive reopen-and-rewrite.

Same-workbook hidden/protected technical storage is the current architecture.

Retention is indefinite by default.

Required therapist history includes at least:

- workload history,
- occupied timeslots,
- active sessions,
- attendance/presence,
- dated absences,
- structured absence reason,
- replacements accepted/covered,
- utilization-relevant cancellations/patient absences,
- robotic/infectious workload where applicable.

Required patient history includes at least:

- date entered system,
- elapsed stay,
- treatment/service days,
- scheduled/performed sessions,
- missed/lost treatments with structured reason,
- treatment timeline,
- meaningful therapist/treatment changes,
- selectable-period trends/totals,
- progress milestones.

Patient links from MASTER are intended to open the patient record/folder and the progress/evaluation view.

---

## 8. Replacement engine: what is proven and must not be thrown away

The replacement/patient-centric engine was built earlier than the current recovery work and historically produced correct operational THERAPIST_DAILY behavior.

Important historical commits include:

- `9f0a11433bc2dd6e6d2d692dee9ed737baa5401d` — patient-centric timetable preview
- `459f6b83f10b4e80cdb7eeb7035d290e1843acdb` — replacement timeslot recommendation layer
- `faaceb1bc6eeedc3833624b77c2a7ad90c1d979a` — guided replacement workflow
- `1074bd1a62746dc6c532787add448f93c2f19f16` — DAILY_INPUT reader/applier
- `16cc7938ef046fde7cba453eb287236d7fdbea3a` — sequential therapist replacement workflow
- `e47c6d46739bfc5713ceab1067832fa205c7d9ac` — compact replacement overlay/source-note behavior

The patient-centric renderer can:

- rebuild affected THERAPIST_DAILY cells from patient-owned assignments,
- preserve existing recurring patients,
- mark/alter only the affected patient line,
- render replacement source information,
- append source note such as `→ therapist time`,
- insert the patient into the replacement therapist/time destination,
- preserve important visual roles.

As of 2026-11-01 the current `Python/tools/prepare_daily_input_in_place.py` STILL calls the same `build_patient_centric_preview_plan(...)` engine.

Therefore the business/rendering engine is not missing from the current branch.

---

## 9. Critical proof from 2026-11-01: old engine still works on current U

This is the most important new forensic result and must not be lost.

The user ran the older copy-based route on the same current U workbook:

Input:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_U.xlsm`

Diagnostic output:

`Excel/previews/DIAGNOSTIC_OLD_ENGINE_FROM_U_20261101.xlsm`

Command path:

`Python/tools/apply_daily_input.py`

Result:

- `DAILY INPUT APPLIED PREVIEW OK`
- Date: `2026-10-30`
- Therapist rows read: `0`
- Patient rows read: `1`
- Operational absences: `1`
- Session cancellations: `0`
- Changed sessions: `1`
- `patient_absent: planner:4:ΦΘ@2026-10-30 [B7]`
- `Input unchanged: True`
- `VBA preserved: True`
- source SHA256 unchanged before/after

This proves that the current patient-centric/DAILY_INPUT planning engine can still read the same current workbook and produce a valid changed copy.

Therefore:

**Do not blame or rewrite the replacement engine unless future evidence contradicts this proof.**

The failure is localized downstream, in the newer in-place VBA application/verification/rollback layer.

Checkpoint commit for this proof:

`f6e77bd89f3dd2ed3b0ca79b275d7d265510f382`

---

## 10. Current root problem: in-place VBA transaction rejects a correct plan

The current architecture prepares a correct Python payload/write plan, then VBA applies it in the same workbook and performs semantic read-back verification before commit/save.

The native failure repeatedly occurs during that VBA transaction.

Current failure evidence from T/U lineage:

`%TEMP%\eka_daily_input_last_error.log`

showed:

- phase: `verify patch`
- patch_index: `2`
- sheet: `MASTER_SCHEDULE`
- cell: `D8` (or D9 in another run depending on affected row)
- error: `DAILY_INPUT read-back mismatch`

Restore then fails secondarily because it uses the same validated patch engine:

`%TEMP%\eka_daily_restore_last_error.log`

- phase: `apply restore payload`
- error: `Restore patch apply failed`

Interpretation:

The transaction applies some visible formatting, then its verifier rejects the MASTER treatment presentation patch, and the whole transaction rolls back before later effects become visible.

This rollback explains why the user reports that nothing else changes:

- no durable patient-line change,
- no durable THERAPIST_DAILY source/destination information,
- no durable ledger/restore progression,
- no useful final operational state.

Do not interpret the absence of these downstream effects as proof that their business logic is broken.

---

## 11. MASTER strike/read-back investigation already performed

The affected MASTER treatment cell was diagnosed directly.

For a failing treatment cell such as D9, diagnostic evidence showed:

- value/text was present,
- whole-cell `Font.Strikethrough=True`,
- whole-cell red color correct,
- Excel COM `Characters(...)` access failed with “member not found”.

This proved that visually correct formatting can exist while COM character-level read-back is unavailable.

Several bounded verifier fixes were attempted:

- whole-cell-first strike verification,
- character fallback only when necessary,
- tri-state handling for `master_daily_non_active_overlay`.

Focused source gate reached:

`63 passed, 2 warnings in 3.88s`

The warnings are known Python 3.14 pytest/ZipFile cleanup warnings and are not functional blockers.

Despite these source gates, native U still produced the same operational result, so further blind verifier patching is not justified.

The next step must instrument the exact sub-checks inside the VBA verifier to identify precisely which property causes patch 2 to fail at runtime.

---

## 12. Exact next technical step

Do NOT build another V/W workbook yet.

Do NOT modify replacement logic.

Do NOT modify THERAPIST_DAILY rendering logic.

Do NOT continue alphabetic rebuilds without new evidence.

The next engineering task is:

**Instrument the in-place VBA verification path for the failing MASTER patch so it logs each individual verification predicate before rollback.**

For patch 2 / `MASTER_SCHEDULE D8` (or equivalent affected treatment cell), capture at least:

- patch intent,
- expected source tag,
- actual cell value/text,
- expected/actual strike state,
- expected/actual font color,
- text-run presence/absence,
- wrap state if relevant,
- fill/border role if relevant,
- result of generic strike verifier,
- result of source-tag presentation verifier,
- which exact `Exit Function` path is taken.

Prefer a single diagnostic log in `%TEMP%` with explicit labels.

The goal is to move from “read-back mismatch” to one exact failing predicate.

Only after that evidence should source code be changed.

---

## 13. Current artifacts and freeze rules

Native artifacts are evidence, not reusable sandboxes.

Important sequence:

- M exposed rich-text status issue, destructive ledger clearing, wrong restore apply call.
- N/O/P/Q were post-build recovery-wrapper failures and that approach is retired.
- R built but failed native MASTER read-back.
- S built from canonical source but still failed at MASTER read-back.
- T built clean but behaved the same.
- U built after bounded tri-state verifier source work but still behaved the same.

Freeze each prior artifact. Never repair them in-place.

The post-build recovery patch/wrapper approach is retired and must not be revived unless new evidence demands it.

Retired tools remain forensic evidence:

- `Python/tools/patch_native_restore_recovery.py`
- `Python/tools/build_unified_preview_native_recovery.py`

Normal builds use only:

`Python/tools/build_unified_preview.py`

---

## 14. Ledger/restore bugs already found and fixed in canonical source

Earlier native M exposed three real issues in the DAILY change ledger/restore layer:

1. destructive THERAPIST_DAILY clearing based only on column A,
2. wrong/noncanonical `ApplyPayloadFileInPlace(...)` call,
3. missing `Dim changeId As String` risk under `Option Explicit`.

Canonical source fixes were made so the ledger should:

- use safe whole-sheet anchoring,
- clear only recognized old ledger blocks,
- avoid broad `A:G + 200 rows` clearing,
- call the canonical validated daily-input payload path,
- declare change IDs correctly.

Do not reopen these as primary root causes while the transaction still dies at MASTER patch 2 before reaching them.

They still require native proof after the MASTER blocker is removed.

---

## 15. Status/reason presentation contract

For patient absence/cancellation on MASTER:

- treatment remains visible,
- treatment is red and struck through,
- patient status/reason is red,
- patient status/reason must NOT be struck through.

Earlier native runs showed the status text incorrectly inheriting strike-through. Canonical source now attempts normalization after `ApplyMasterView`.

This behavior has not yet received final native green proof because the transaction rolls back earlier.

---

## 16. Restore contract to prove after Apply becomes green

Once a patient absence Apply completes successfully, the next native proof is only the same-row restore.

Expected restore behavior:

- no VBA editor interruption,
- restore succeeds in the same workbook,
- DAILY_INPUT source A:E for that change is cleared appropriately,
- MASTER returns to normal/present state,
- THERAPIST_DAILY base state returns,
- pre-existing operational details remain untouched,
- ledger row becomes `Επαναφέρθηκε`,
- workbook stays open.

Do not run `Restore all`, therapist restore, or `CloseOperationalDay` until same-row restore is proven.

---

## 17. Close-day/history implementation status

Pure Python history/finalization architecture has significant green coverage and should be preserved.

Implemented areas include:

- day finalization model,
- therapist day history,
- finalized snapshots,
- append-only events,
- snapshot versioning concepts,
- close-day orchestration,
- same-workbook rollback-oriented VBA architecture.

Close-day native progress is intentionally blocked until DAILY_INPUT apply/restore is reliable.

Do not jump forward into statistics UI while daily truth is still operationally unstable.

---

## 18. Important recovery/checkpoint documents

Key documents include:

- `Documentation/ONBOARDING.md`
- `Documentation/RECOVERY_FORENSIC_CHECKPOINT_2026-10-30.md`
- `Documentation/RECOVERY_APPLICATION_BUILD_GATE.md`
- `Documentation/RECOVERY_IN_SESSION_GATE_2026-10-30.md`
- `Documentation/RECOVERY_OPERATOR_GATE_2026-10-31.md`
- `Documentation/RECOVERY_H_PATIENT_ABSENCE_AND_RESTORE_CONTRACT_2026-10-31.md`
- `Documentation/RECOVERY_DAILY_CHANGE_RESTORE_2026-10-31.md`
- `Documentation/RECOVERY_DAILY_CHANGE_LEDGER_IN_PLACE_2026-10-31.md`
- `Documentation/RECOVERY_J_TO_M_HANDOFF_2026-10-31.md`
- `Documentation/RECOVERY_CANONICAL_SOURCE_GREEN_GATE_2026-11-01.md`
- `Documentation/RECOVERY_NATIVE_S_BUILD_2026-11-01.md`
- `Documentation/RECOVERY_NATIVE_S_FAILURE_2026-11-01.md`
- `Documentation/RECOVERY_NATIVE_T_BUILD_2026-11-01.md`
- `Documentation/RECOVERY_NATIVE_T_FAILURE_2026-11-01.md`
- later 2026-11-01 continuity/forensic checkpoints including the old-engine isolation proof.

When evidence conflicts, prefer the newest dated checkpoint plus this master continuity document.

---

## 19. Important recent commits

Examples of high-value recent commits/checkpoints:

- `1b7156f` — whole-cell-first MASTER strike read-back fallback source commit
- `09870c3` — corrected source migrator syntax
- `0caabbe` — updated stale regression for bounded tri-state contract
- `f6e77bd89f3dd2ed3b0ca79b275d7d265510f382` — old-engine diagnostic proof checkpoint

Earlier key recovery checkpoint commits include:

- `2f605ba66a6608e1b6e55e32515a6acec2eb763f` — forensic checkpoint
- `d7f8232fe2334bbedd2c5adedb1787ddcd420b25` — operator gate
- `663bdce8fb22b4ee51e45f6959080443c93e6d32` — patient absence/restore contract
- `8c5932a13b5ded249a436dc552ce7899d87f9b97` — ledger/restore checkpoint
- `5eda4961360f185891837fc09104462d8cacfe67` — green pre-I checkpoint
- `e6f957294b380194ea3a93f27ecaaa349eb17ebd` — zero-state restore fix
- `bb785772c68b51bc2adacdf8001fef7e70cc67e5` — layout-aware MASTER status normalization
- `f03c32215a4f1eb2d8b6aa9a12d77e57b03d98aa` — regression for that normalization

Do not rely on commit IDs alone. Read the associated checkpoint/document when resuming.

---

## 20. Known warnings that are not blockers

Python 3.14 can emit `_pytest.unraisableexception` / `ZipFile.__del__` warnings such as:

`ValueError: I/O operation on closed file`

These have appeared in otherwise green test gates.

Treat them as known cleanup warnings unless they become associated with an actual functional failure.

---

## 21. What must NOT be repeated

Do not:

- continue blind U/V/W/X workbook generation,
- revive the retired post-build recovery wrapper,
- assume THERAPIST_DAILY logic is broken merely because rollback hides its result,
- rewrite the replacement solver without evidence,
- clear broad THERAPIST_DAILY regions,
- repair failed smoke workbooks,
- use production as a test surface,
- treat `4a2eed0` as the historical replacement-engine milestone,
- infer that a successful Python test alone proves native Excel behavior,
- infer that a native visual failure means the Python planner is wrong when the transaction log shows rollback earlier.

---

## 22. Exact resume instructions for a new conversation

A new ChatGPT conversation should do this in order:

1. Read `Documentation/ONBOARDING.md`.
2. Read this file: `Documentation/PROJECT_CONTINUITY_MASTER_2026-11-01.md`.
3. Read the latest `Documentation/CURRENT_HANDOFF.md`.
4. Inspect the newest recovery checkpoint(s) from 2026-11-01.
5. Confirm branch is `work/history-statistics-rebuild`.
6. Do not build another preview immediately.
7. Start from the proven fact that the copy-based old engine works on current U.
8. Instrument the VBA in-place verifier for the failing MASTER patch and identify the exact predicate causing rollback.
9. Make the smallest canonical source fix supported by that evidence.
10. Add/update regression tests.
11. Run focused source gate.
12. Build one fresh unique native artifact only after the source gate is green.
13. Run one narrow patient-absence Apply smoke.
14. If Apply succeeds, run same-row Restore only.
15. Update checkpoint documentation with every result.

---

## 23. Final continuity principle

The project should advance by preserving proven layers and isolating broken layers.

The accumulated evidence currently says:

- the core scheduling/patient-centric engine is valuable and still functional,
- the current in-place transaction verifier is the active failure surface,
- history/statistics remain the next major product milestone after daily operational truth is stable,
- the center’s real operational needs, not the mechanics of Excel, define success.

When in doubt, protect the working engine, preserve the immutable baseline, use diagnostics before edits, and return to the five project goals at the top of this document.
