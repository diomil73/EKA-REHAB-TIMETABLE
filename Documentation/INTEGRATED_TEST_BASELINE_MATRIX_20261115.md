# DIOMIL integrated test baseline and acceptance matrix

Date: 2026-11-15
Branch reviewed: `work/afternoon-medical-scope-20261115`

## Purpose and interpretation

This document distinguishes (1) mechanisms with operator-confirmed native Excel evidence, (2) mechanisms with source-only or partial native evidence, and (3) planned capabilities that are not yet implemented/accepted. It is the selection guide for building **one complete synthetic integrated DIOMIL test workbook**. A workbook filename, a green pytest suite, or an isolated preview is never enough to establish whole-workbook acceptance.

Status terms:
- **NATIVE ACCEPTED / FROZEN**: documented operator-confirmed Excel workflow on that exact artifact; do not modify or reuse as a writable candidate.
- **NATIVE GATES GREEN; MANUAL ACCEPTANCE PENDING**: installer and read-only persisted audits passed; user-visible interaction is not fully accepted.
- **SOURCE GREEN ONLY**: Python/source tests passed; not native Excel proof.
- **DESIGN / NOT IMPLEMENTED**: documented requirement without an accepted implementation.
- **BLOCKED**: safety, provenance, or business-rule gate prevents the next action.

## 1. Operational mechanisms with documented native evidence

| Capability | Canonical evidence | What is actually evidenced | Boundary |
|---|---|---|---|
| Daily patient absence Apply / Restore | `UNIFIED_CLOSE_DAY_RECOVERY_20261101_X.xlsm` | Native Apply once / Restore once for the documented all-day patient-absence scenario; MASTER and THERAPIST_DAILY removal/restoration, formula/status restoration and ledger ACTIVE→RESTORED were verified. | This proves the tested scenario, not every absence combination or all future recurring edits. |
| Close/finalize day (H2) | `UNIFIED_CLOSE_DAY_VERIFIED_PREVIEW_20261025_Z.xlsm` (see H2 runbook / handoff) | Native finalized snapshot, therapist/day history and append-only finalization event. | Do not reuse the frozen artifact. |
| Day amendment/version chain (H3) | `UNIFIED_CLOSE_DAY_RECOVERY_20261101_AB.xlsm` (see H3 handoff) | Native v1→v2 amendment, superseded/current-final version chain and persisted readback. | Do not reuse the frozen artifact. |
| Therapist historical report (H4) | `UNIFIED_THERAPIST_HISTORY_UI_20261102_AQ.xlsm` | Native daily, weekly, monthly and overall report modes; button row remains visible through mode changes; CURRENT_FINAL-only aggregation. | AQ is frozen. |
| Patient-session history (H5) | `UNIFIED_PATIENT_HISTORY_H5_20261103_AX.xlsm`, `..._AY.xlsm` | Native first close and v1→v2 amendment; 80 patient-session history rows and version-aware readback for the recorded scenario. | AX/AY are frozen. |
| Contextual day-save UX (H6) | `UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm`, `..._BB.xlsm` | Native contextual Save Day flow and the H6 baseline used by later H7/H8 installs. | AZ is the trusted frozen source for the H7/H8 chain; BB's exact scope remains in its checkpoint. Never overwrite. |
| Integrated history toolbar and reports (H7) | `UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm` | Operator-confirmed integrated toolbar, RefreshAll, patient and therapist report entry points, modes and return navigation. | Functional acceptance did not mean full analytics UX acceptance; later H8 work refined it. BF frozen. |
| H8 patient/therapist dashboards | `UNIFIED_PATIENT_HISTORY_H8_20261107_BL.xlsm` | Operator-confirmed H8 UI scope: patient overview and selected-patient chart categories (scheduled/performed/absences), active period styling, sticky chart while scrolling, therapist panels, report navigation/protection and single terminal RefreshAll success. | BL is frozen. This is the accepted scope tested then, not completion of every analytics metric or 15-day patient progress. |
| Synthetic finalized history dataset | `UNIFIED_H8_DEMO_HISTORY_20261108_BO.xlsm` | Independent read-only cross-sheet proof: 25 service dates, 522 patient-session rows, 150 therapist-day rows, 150 synthetic attendance facts, 25 finalization events; unchanged by verifier. | This proves the synthetic history ledger and references against its registry. It does **not** certify every roster/patient-bearing sheet in the whole workbook as synthetic. BO is immutable/frozen. |

The earlier H7 protection follow-up BI had native install/protection/UI structural audits, but its analytics UX was not accepted; it is not a replacement for BL's later H8 acceptance.

## 2. Therapist H10 dashboard and management presentation: partial / not fully accepted

- `UNIFIED_H10_DEMO_THERAPIST_20261108_BU.xlsm`: native visual evidence for an earlier therapist dashboard iteration; preserve frozen. Do not infer that every later H10 feature is accepted from BU.
- `UNIFIED_H10_DEMO_THERAPIST_20261108_BX.xlsm`: native installation and read-only history/protection/integrated-UI gates were operator-confirmed green. The user then reported concrete visual defects (date-formatted aggregate count, stale traffic fills) and requested redesign of internal and standalone management views. Therefore BX is **not** the final visual acceptance for those later requirements.
- Later BY source changes include management layout, four macro-free mode-navigation buttons in the independent XLSX, KPI cards and privacy/readback checks. The documented source/tests/native acceptance for BY were not confirmed green in the latest evidence reviewed.
- H10's authentic therapist absence/working-day chart remains unavailable because the finalized source does not yet persist reliable dated attendance/absence and working-day facts. The UI must label this unavailable, never display a fabricated zero.
- The separate management-only XLSX export is a distinct artifact and needs its own fresh native export, privacy/readback and visual approval. Never share the operational XLSM.

Consult `H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md`, `H10_BV_NATIVE_RUNBOOK_20261108.md`, `H10_BW_NATIVE_RUNBOOK_20261108.md` and the latest H10 appendices in `CURRENT_HANDOFF.md` for exact candidate chronology. Historical failed candidates stay frozen and must never be retried.

## 3. Patient statistics and the 15-day progress/evaluation view

### Patient statistics already implemented/accepted in a defined scope

- Patient session history is persisted from finalized day facts, version-aware and read-only for reports.
- Patient report modes and chart views were incorporated into the H7/H8 chain.
- BL's accepted H8 scope includes scheduled, performed and absence categories in overview and selected-patient charts.
- The H8 design specifies patient identity by stable PatientID joined to a display name; participation percentage must use a documented eligible-session denominator and handle zero denominator as unavailable, not zero or 100%.

### 15-day patient progress/evaluation: DESIGN ONLY

The specialist source of truth `STATISTICS_HISTORY_DESIGN.md` specifies:
- MASTER's physiotherapy-time cell opens the patient's progress/evaluation view.
- Manual evaluation approximately every 15 days; append-only historical milestones.
- Each milestone records the examined question/goal, `Από` date, `Έως` date and manually selected outcome.
- Fixed outcome meaning: green = positive/progress, grey = neutral/no meaningful progress, red = deterioration.
- PatientID-based history, chronological trend/sparkline, unaffected by later therapist/time/cell changes.
- Additional planned views: current inpatients with room/name/current 15-day goal/latest evaluation/history/trend; overdue 15-day evaluations; treatment-type timeline; lost treatments and reasons.

**No native-accepted patient 15-day progress/evaluation workflow is recorded in the evidence reviewed.** Do not represent it as already built. It is a distinct remaining product slice after the existing H5/H8 session-history foundation.

## 4. Other explicit open mechanism / release boundaries

The user's recollection that most day-to-day mechanisms are already built is broadly consistent with the substantial H2-H8 native evidence. However, repository records do **not** support the claim that every mechanism is complete:

1. **Global multi-therapist-absence replacement optimizer** remains open in `PROJECT_STATUS.md`; existing single-patient candidate ranking is not the same as solving all displaced patients together.
2. **Permanent recurring therapy change write / rollback / operational UI** remains incomplete in `AUTHORITATIVE_RESUME_20261112.md`: stable AssignmentIDs were natively migrated on a synthetic copy and an event reader/append plan exists, but no native permanent-event append, authoritative transaction, rollback or operational VBA UI has been proven.
3. **Real-source import/admission** remains blocked by unresolved identity/source holds/provider conflicts. The trial admission and weekly projection tools are planning/preflight, not a native importer.
4. **Final worksheet protection hardening** is not a blanket completed mechanism; write surfaces and approved manual-edit zones require classification before protection.
5. **Afternoon physiotherapy integration** on the active branch has a read-only selected-patient display and source-green tests (latest reported five-file gate: 49 passed in 9.22s). It has no afternoon save path, no native Excel acceptance and no complete whole-workbook synthetic source yet. Its conflict namespace is afternoon-only, separate from the morning programme.
6. **Whole-workbook synthetic fixture** has not yet been certified. The tracked populated `Rehab_Center_System_v27_1.xlsm` is excluded. H8/H10 synthetic-history proofs do not establish that all patient/staff rosters and every patient-bearing sheet are synthetic.

These are separate from the core workflows already natively demonstrated. Keep the product status granular rather than marking all mechanisms either complete or incomplete.

## 5. Selection strategy for one complete synthetic DIOMIL workbook

The target is one integrated XLSM containing the full application shell and all already-integrated workflows, with synthetic patient/staff identities and consistent IDs across every patient-bearing sheet. It must not be assembled by simply copying one frozen preview or by treating a preview filename as a specification.

1. Use the native-accepted artifacts above as **evidence and source-component references**, not as writable build targets.
2. Audit the latest compatible application baseline and all VBA/Python installers, forms, buttons, named ranges, formulas, validations, protection and hidden/veryHidden sheets.
3. Inventory all patient/staff data-bearing surfaces (including hidden/history/weekly/outpatient/pool/robotic/psychology/attendance/assignment/event stores and embedded names/caches).
4. Build a deterministic synthetic cohort and synthetic scenarios, preserving the full application structure and compatible components. Prove no real patient/staff identity or clinical note survives anywhere in the workbook.
5. Reconcile every cross-sheet PatientID, TherapistID, StudentID, AssignmentID, EventID and session reference.
6. Run full source regression and dedicated synthetic provenance audits.
7. Create a fresh, never-existing native Excel candidate once; save, close and independently reopen read-only; compare required workbook structures, formulas, VBA and protected/hidden states.
8. Run the full manual end-to-end acceptance sequence. Freeze all candidates after the attempt, success or failure; never overwrite or repair in place.

No frozen acceptance artifact, trusted synthetic history source or production workbook may be modified as a shortcut.

## 6. Last known source test result

The latest operator-confirmed result in the active afternoon branch at the time of this document is **49 passed in 9.22s** across the five targeted afternoon source modules. This is source-only evidence for that branch slice. It is not a full-project test run, a native Excel save/reopen proof, or whole-program acceptance.

For a verbose or lengthy test run, save the complete console output to a local text file and attach that file to the conversation; a short pytest summary is enough when all tests pass. If failures occur, preserve the first full traceback and the command that produced it.

## 7. Read-first references

- `Documentation/CURRENT_HANDOFF.md`
- `Documentation/PROJECT_STATUS.md`
- `Documentation/STATISTICS_HISTORY_DESIGN.md`
- `Documentation/H8_STATISTICS_DASHBOARD_OPERATOR_FEEDBACK_2026-11-07.md`
- `Documentation/H10_THERAPIST_CHARTS_USER_REQUIREMENTS_20261108.md`
- `Documentation/AFTERNOON_PHYSIOTHERAPY_MEDICAL_DEFERRAL_POLICY_20261115.md`
- `Documentation/AUTHORITATIVE_RESUME_20261112.md`


## 2026-11-15 — Full repository regression: 1,806 passed; stale VBA-verifier fixture corrected

Operator uploaded the complete Windows pytest log `DIOMIL_TEST_FULL_20261115_161242.txt`. It records:
- 1,808 tests collected.
- **1,806 passed, 1 failed, 1 skipped** in 36.72s.
- 27 warnings.
- Skip: `Tests/test_afternoon_sheet_preview.py:68`, symlinks unavailable on this Windows environment.
- Sole failure: `Tests/test_patient_registration_scheduling_slice.py::test_patient_program_vba_verification_accepts_only_new_runtime_and_form`.

### Root cause

The implementation's `_verify_patient_program_vba()` correctly enforces the integrated afternoon read-only panel contract: form markers `RefreshAfternoonReadOnlyPanel`, `lblAfternoonState`, `Me.Height = 665`; runtime markers `LoadPatientAfternoonState`, `afternoon-state-vba`; and the actual `lblAfternoonState` designer control. The failing test's fake VBA components only represented the older morning-only form/runtime and had no Designer.Controls implementation. Thus the test fixture was stale after the PM panel was integrated; the log does not show an application runtime failure or a native Excel failure.

### Correction committed, not yet rerun

- `dc4e331`: updated the synthetic fake components to include the current PM form/runtime markers and designer label; added a negative regression proving that a missing designer label still fails closed.
- `f30c32b`: clarified the verifier comment to describe the PM panel as part of the integrated form contract.

No application installer logic or patient/workbook data was changed. No Excel workbook was opened or modified by this test-fixture correction. The full suite remains **not yet green** until the operator reruns it.

### Next commands

First run the focused test:

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_patient_registration_scheduling_slice.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Patient registration scheduling-slice test failed" }
```

If green, rerun the complete suite and preserve the full output to TXT only because this is a large result:

```powershell
$log = Join-Path $PWD ("DIOMIL_FULL_REGRESSION_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
python -m pytest -ra --tb=long *> $log
$testExitCode = $LASTEXITCODE
"EXIT CODE: $testExitCode" | Add-Content -LiteralPath $log
Get-Item -LiteralPath $log | Select-Object FullName, Length
Get-Content -LiteralPath $log -Tail 25
if ($testExitCode -ne 0) { throw "SAFETY STOP: Full regression suite failed; inspect the TXT report" }
```

The `*> $log` redirection captures PowerShell success/error streams without Transcript. Capture `$LASTEXITCODE` immediately after pytest, before running other commands.


## 2026-11-15 — FULL REPOSITORY REGRESSION GREEN AFTER FIXTURE UPDATE

Operator reran the full suite after pulling the focused-test fixture correction.

- Full command: `python -m pytest -ra --tb=long`
- **1,808 passed, 1 skipped, 0 failed** in **30.54s**.
- **EXIT CODE: 0**.
- Skip: `Tests/test_afternoon_sheet_preview.py:68`, because symlinks are unavailable in the current Windows environment.
- 26 warnings were reported. They include Python 3.14/openpyxl ZipFile deallocator warnings and an openpyxl Data Validation extension warning in the existing daily-input baseline audit. They did not fail the suite and should be assessed separately; do not perform broad cleanup or modify workbook files merely to silence warnings.

The previous section titled “Correction committed, not yet rerun” is now superseded: the focused test passed **8 passed in 2.66s**, and the complete suite is now green. The stale fake VBA component fixture was updated to include the integrated afternoon panel contract and a negative test for a missing designer label.

This is a **source regression GREEN** checkpoint only. It is not native Excel save/reopen proof, full integrated workbook provenance proof, nor manual end-to-end acceptance. No Excel workbook was opened or modified during this source-test correction and regression run.

### Next safe project phase

Proceed with the **full DIOMIL synthetic integration workbook** plan, not an afternoon-only preview. The immediate engineering task is a read-only architecture/data-surface audit of the current compatible application baseline: identify the exact workbook shell/component source and inventory every patient/staff-bearing sheet, hidden/veryHidden surface, named range, formula/cache, VBA component and linked identity/assignment/history table. Then design a deterministic synthetic roster and prove full-workbook synthetic provenance before any native installer or Excel write operation. Keep the accepted/frozen H2-H8 artifacts and populated operational workbook immutable.
