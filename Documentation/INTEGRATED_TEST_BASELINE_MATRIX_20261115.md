# DIOMIL integrated test baseline and acceptance matrix

## Verified code gate 2026-11-16: dual-source inpatient/outpatient demo (not real operator acceptance)

At SHA `70e1b88f6e3289bbc79fe411d4724e327622ff7f` GitHub Actions [43809724956](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/43809724956) passed 2,029 Python tests on Windows, 2,029 on Linux, 4/4 jobs. The new source-only outpatient projector and separate source collision/weekday/native-writer/readback guards are covered by fictional tests. Its preceding candidate had one invalid test fixture (missing standalone outpatient file), corrected before this clean gate. The full build now requires a separate approved `--outpatients` XLSX; outpatient records are not generated from the inpatient A:V manual. No native real patient dataset was read, no private native XLSM was built, no interactive user smoke was performed; release and operator acceptance remain blocked on fresh source approvals, F/U/V semantics, source parity and native Save/Close/Reopen. Keep H2-H10 historical proof frozen.


## 2026-11-16 — Executable private real-source DEMO preflight

**New executable interface:** `Python/tools/prepare_real_data_practice_demo.py` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`. Default operation is **read-only full real-source validation** for an authorized current A:V manual XLSX, confirmed review XLSX, entirely empty application-ready VBA XLSM and independently approved private `clinical_approvals.json`. Uses the strict production source alignment, all-patient approved ID check, original weekday/time/provider reconciliation, unresolved F/U/V stop gates, existing native shell checks and SHA-256 stability of the clinical-approval JSON. Every input must remain off Git checkout. No real patient name, ID, room, doctor, MRN, source path or clinical exception content is emitted to stdout/stderr on handled errors. `--build --output <PRIVATE_*_DEMO_*.xlsm>` is an *explicit* Windows + native Excel/COM, exclusive-new-file second step and repeats the same approval gates. **No real source file has been run; no real-data XLSM exists, and end-to-end practical acceptance is still pending.** Previous synthetic/H8 workbook results do not count as the operator's comparison. First actual action is an authorized private preflight with current manual and approved clinical review, not a random synthetic seed. Native failure or changed approvals after a native operation requires freezing the candidate, not treating it as accepted. See the dedicated operator CLI runbook for exact commands/JSON schema.

**Current blocker:** F/U/V and independently sourced outpatient bookings remain unavailable to the current *inpatient-only* A:V admission bridge. If occupied, the complete real-data source fails closed, never silently drops treatments. Do not alter frozen H2-H10 evidence. Update all living handoff/status/test and delivery documents again with actual source-specific counts and results **only after** the first private native execution, without publishing PII.

## ΕΠΙΧΕΙΡΗΣΙΑΚΗ ΕΝΗΜΕΡΩΣΗ 2026-11-16: ΠΡΑΓΜΑΤΙΚΟ DEMO

**Supersedes previous phrase “one complete synthetic integrated DIOMIL test workbook” as end-user acceptance.** We retain a synthetic integration workbook ONLY as a technical fixture. **User evaluation requires a private native real-data XLSM** with the actual patients, every weekday/time/treatment/provider from the current manual schedule, not random fixtures. Tests GREEN (1.960 Windows/Linux, run 43769366651) are **not** real-data/operator acceptance. Native/real-source/manual side-by-side/status of F/U/V/user flows remain **PENDING**. All acceptance rows must distinguish code/synthetic, native seeded, full source match, operational interaction, save/restore, and signed-off pilot. Full definitive matrix and handoff: `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`.

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


## 8. Repository workbook-builder boundary discovered after the full regression

Read-only source inspection of `Python/tools/build_unified_preview.py` and `Python/rehab_excel/application_contract.py` confirms that the existing `UNIFIED_APP_SMOKE` builder is **not** the requested complete DIOMIL test workbook. Its documented build sequence combines registration menu/forms, patient registration bridge, patient edit flow, outpatient schedule form, outpatient conflict guard, DAILY_INPUT, MASTER projection/navigation/visual styling. Its application contract checks a specific set of 25 VBA components plus `OUTPATIENT_SCHEDULE` and `DAILY_CHANGE_LEDGER`; it does not by itself prove integration of the complete H2-H8 history/dashboard workflows, all clinical scheduling mechanisms, all management exports, or the 15-day patient progress feature.

The repository currently tracks only one binary Excel workbook: `Rehab_Center_System_v27_1.xlsm` (309,543 bytes in the reviewed tree). It is a populated clinical baseline and remains excluded from all synthetic fixture writes. There are no tracked XLSM/XLSX files under `Excel/` in this branch. The accepted/frozen previews referenced in the checkpoint matrix are local artifacts, not tracked repository files.

Therefore:
- Do not run `build_unified_preview.py` against the tracked baseline as a shortcut to the full synthetic test artifact.
- Do not treat `UNIFIED_APP_SMOKE` or the H8 demo as the full application test workbook.
- Next selection step is a read-only local manifest of all existing XLSM/XLSX candidates (full path, exact filename, size, last-write timestamp). Then cross-reference candidates with the native acceptance matrix and provenance evidence; do not open or modify any candidate during the manifest step.


## 2026-11-15 — Workbook manifest inspected; H10 candidate triage, no workbook selected

The operator supplied and the project reviewed `DIOMIL_WORKBOOK_MANIFEST_20261115_162813.txt`.

- Manifest total: **337 workbooks** = **336 .xlsm + 1 .xlsx**.
- The inventory is metadata-only: full paths, names, sizes and last-write times. The manifest explicitly states **NO WORKBOOK OPENED OR MODIFIED**.
- The list contains multiple generations (H5–H10), history/demo artifacts, older master/tool-bar smoke workbooks, and forensic recovery-lineage files.
- Most recent named H10 artifacts in the manifest:
  - `Excel/previews/UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm` — 629,797 bytes; 12 Nov 2026 23:40:58.
  - `Excel/previews/UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm` — 622,070 bytes; 12 Nov 2026 23:16:38.
  - `Excel/previews/UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` — 731,415 bytes; 12 Nov 2026 23:10:08.
  - `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` — 731,415 bytes; 11 Nov 2026 01:14:02.
- `UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm` is the newest H10-named candidate by manifest timestamp, **not an approved integration baseline**. Filename, size and timestamp do not prove workbook content, synthetic-only provenance, completeness, or acceptance status.
- No workbook has been selected, opened, modified, migrated, or designated as the complete synthetic integration base. Preserve the populated clinical baseline and all accepted/frozen artifacts.

Regression evidence must be kept distinct:
- The user-provided terminal tail reports **1808 passed, 1 skipped, 26 warnings in 30.54s; EXIT CODE: 0**.
- The separately uploaded `DIOMIL_TEST_FULL_20261115_161242.txt` records an earlier/different run with **1 failed, 1806 passed, 1 skipped, 27 warnings in 36.72s**. Its failure is `Tests/test_patient_registration_scheduling_slice.py::test_patient_program_vba_verification_accepts_only_new_runtime_and_form`, reporting missing afternoon form/runtime markers and the read-only label control.
- These are two distinct artifacts/results; do not merge them or claim the uploaded failing report is green. The exact full path/name and complete contents of the successful `DIOMIL_FULL_REGRESSION_...` report still need to be matched to the supplied green terminal tail before it is treated as the authoritative full-regression artifact.

Next: finish metadata classification and cross-check the three leading H10 candidates against existing acceptance/provenance documentation. Only then nominate a candidate for a controlled, read-only content inspection. Do not perform any workbook write, native Excel open/save, operational migration, or patient-data action as part of this selection stage.


Cross-reference to the pre-existing acceptance matrix: the whole-workbook synthetic fixture is still **NOT CERTIFIED**; the tracked `Rehab_Center_System_v27_1.xlsm` is a populated clinical baseline and is explicitly excluded. Existing H8/H10 synthetic-history verification covers only its documented history scope and does not establish that every registry, schedule, hidden sheet, cached value, identity or embedded source in a candidate is synthetic. The H10 dashboard lineage is documented as partial/not fully accepted. Therefore none of the four newest H10 candidates may be promoted to a whole-program synthetic source solely from its name, timestamp, or H8/H10 history evidence.

Specific manifest traps and references:
- `UNIFIED_H10_DEMO_SYNTHETIC_ONLY_20261109_CB.xlsm` must **not** be treated as a fully synthetic whole-workbook fixture. The documented helper isolates synthetic history-ledger rows but does not sanitize the PATIENTS registry and all other patient-bearing sheets.
- `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` belongs to the frozen H10 therapist lineage; preserve it unchanged. The separate `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` is a different file and its name alone does not prove acceptance or suitability.
- `UNIFIED_H8_DEMO_HISTORY_20261109_CA.xlsm` is a history demonstration artifact, not proof of a complete synthetic application roster.
- The sole `.xlsx` entry is `DIOMIL_MANAGEMENT_BY_DEMO_20261108.xlsx`; its filename does not establish that it has passed the independent privacy/export audit required for a shareable management-only deliverable.

## 2026-11-15 — Read-only whole-workbook structure auditor added

Added:
- `Python/tools/audit_integrated_workbook_structure.py`
- `Tests/test_audit_integrated_workbook_structure.py`

Purpose: non-identifying architecture inventory for one or more XLSM/XLSX candidates. It uses ZIP metadata plus two openpyxl read-only passes (formula view and cached-value view). It reports workbook SHA-256 and size, ZIP part categories, VBA project size/hash, defined-name count, worksheet names/visibility/dimensions, non-empty cell counts, formula counts, cached formula-result counts, and broad schema role hints for patient/staff/schedule/history/afternoon/pool/robotics/psychology/registry surfaces. It never prints cell values, formula text, cached values, patient/staff identifiers or clinical notes. It verifies the source file hash and size are unchanged after each audit. Conservative scan caps return exit code 3 and explicitly mark the inventory incomplete rather than silently accepting a partial scan.

This is a **source-level read-only file inspection**, not native Excel COM inspection. It does not enumerate VBA component names/UserForm controls or prove the synthetic provenance of cell values. VBA is currently represented only by presence, byte length and SHA-256 of `xl/vbaProject.bin`; a separate read-only COM component inventory remains a later step.

### Operator execution gate

After pulling the latest branch, run the focused tests first:

```powershell
git pull --ff-only origin work/afternoon-medical-scope-20261115
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
$env:PYTHONPATH = (Join-Path (Get-Location).Path "Python")
python -m pytest Tests/test_audit_integrated_workbook_structure.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Structure auditor tests failed" }
```

If focused tests pass, inspect only the four H10 candidates from the metadata manifest. The script reads files through ZIP/openpyxl read-only APIs; it does not launch Excel or save workbooks. Save this larger output to TXT:

```powershell
$log = Join-Path $PWD ("DIOMIL_H10_STRUCTURE_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
python Python\tools\audit_integrated_workbook_structure.py `
  "Excel\previews\UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm" `
  "Excel\previews\UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm" `
  "Excel\previews\UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm" `
  "Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm" *> $log
$auditExitCode = $LASTEXITCODE
"EXIT CODE: $auditExitCode" | Add-Content -LiteralPath $log
Get-Content -LiteralPath $log -Tail 35
if ($auditExitCode -ne 0) {
    throw "SAFETY STOP: One or more workbook structure audits need review; upload the TXT."
}
```

If any candidate is missing, corrupt, or exceeds a scan cap, the combined run is nonzero by design. Do not change the candidate paths, repair the file, or substitute a different workbook without reviewing the manifest first. The four candidates remain unselected and are not authorized for any write/native installation. The tracked populated clinical baseline and all frozen acceptance artifacts remain excluded.


## 2026-11-15 — Full regression after integrated workbook structure auditor

Operator supplied the complete report `DIOMIL_FULL_REGRESSION_20261115_164649.txt`.

Authoritative result from this exact report:
- Collected: **1,814 tests**
- Passed: **1,813**
- Failed: **0**
- Skipped: **1** — `Tests/test_afternoon_sheet_preview.py:68`, reason: symlinks unavailable on this Windows environment
- Warnings: **26**
- Duration: **39.46s**
- Process exit code: **0**

The report confirms `Tests/test_audit_integrated_workbook_structure.py` ran as part of the full suite and passed all five tests. The separately executed focused run had also reported **5 passed in 2.20s**.

This establishes source-level Python regression GREEN for this revision. It does **not** establish native Excel acceptance, workbook candidate provenance, or whole-application synthetic integration.

The uploaded artifact contains only the full regression output. No `DIOMIL_H10_STRUCTURE_*.txt` report has yet been received, so no conclusion is recorded about the four H10 workbook candidates. They remain unselected and untouched.

Next operator action: run the already documented four-candidate read-only structural audit only (do not rerun the full regression), then upload `DIOMIL_H10_STRUCTURE_*.txt`. Review candidate worksheet visibility, data-bearing surfaces, formulas/cached-result coverage, package parts and VBA project fingerprints before any candidate is nominated. Keep the tracked populated clinical baseline and all frozen acceptance artifacts excluded.


## 2026-11-15 — H10 structure audit invocation stopped by PowerShell stderr policy; retry command corrected

The first four-candidate invocation produced `DIOMIL_H10_STRUCTURE_20261115_164733.txt` containing only `EXIT CODE: -1`. The console shows that PowerShell raised `NativeCommandError` when openpyxl emitted the non-fatal warning `Data Validation extension is not supported and will be removed`. The shell had `$ErrorActionPreference = "Stop"`, so native Python stderr interrupted the invocation. This is **not evidence that the structural auditor found a corrupt workbook or failed its structural checks**; no complete audit output was produced.

The warning is emitted by openpyxl while reading legacy Excel data-validation extensions. The auditor uses read-only workbook loading and never saves the workbook. Do not interpret this warning as a workbook modification.

### Corrected operator command (no tests need rerunning)

Use a fresh report filename and temporarily set `$ErrorActionPreference = "Continue"` only around the native Python command, while keeping `*> $structureLog` to capture stdout and stderr. Restore the previous preference afterwards and read the actual Python `$LASTEXITCODE`:

```powershell
$structureLog = Join-Path $PWD (
    "DIOMIL_H10_STRUCTURE_{0}.txt" -f
    (Get-Date -Format "yyyyMMdd_HHmmss")
)

$previousEAP = $ErrorActionPreference
try {
    $ErrorActionPreference = "Continue"

    python Python\tools\audit_integrated_workbook_structure.py `
      "Excel\previews\UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm" `
      "Excel\previews\UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm" `
      "Excel\previews\UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm" `
      "Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm" *> $structureLog

    $structureCode = $LASTEXITCODE
}
finally {
    $ErrorActionPreference = $previousEAP
}

"PYTHON EXIT CODE: $structureCode" | Add-Content -LiteralPath $structureLog
Get-Content -LiteralPath $structureLog -Tail 45
Write-Host "STRUCTURE REPORT: $structureLog"
Write-Host "PYTHON EXIT CODE: $structureCode"

if ($structureCode -ne 0) {
    throw "SAFETY STOP: Structural audit did not complete cleanly; upload the full TXT."
}
```

Only this read-only structure audit needs to be rerun. Do not rerun regression, do not suppress all Python warnings inside the audit tool, do not open or save the candidates in Excel, and do not change candidate paths. The four workbooks remain unselected and untouched.


## 2026-11-15 — Four-candidate H10 read-only structural audit completed; only final-candidate tail received

Operator reports successful completion of the four-file invocation and supplied the tail of `DIOMIL_H10_STRUCTURE_20261115_165132.txt`:
- Python exit code: **0**
- `ALL CANDIDATE STRUCTURE AUDITS COMPLETE`
- `SOURCE HASH UNCHANGED: YES`
- `NO WORKBOOK WAS MODIFIED`

The visible tail is the final candidate block in the invocation order, namely `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm`. Its visible structural metrics:
- 25 worksheets; 36,883 scanned cells; 5 defined names.
- VBA project: 1,305,600 bytes; SHA-256 `285b7f4aa20a195535d300928f26b2884255797c1817d4f2a7e32ac82b4033b9`.
- Package: 11 drawings, 2 charts, 1 shared-strings part, 2 VML drawings; no external links, connections, query tables, pivots, slicers, ActiveX, embeddings or custom UI reported.
- `PATIENT_PLANNER`: 2,745 formulas, 465 cached formula results.
- `MASTER_SCHEDULE`: 1,021 formulas, 549 cached formula results.
- `STATISTICS`: 43 formulas, 43 cached formula results.
- `PATIENT_HISTORY_REPORT` is empty (`A1:A1`); `OUTPATIENT_SCHEDULE` contains only its header row.
- `SESSIONS` is hidden; `_PY_LISTS`, `DAILY_CHANGE_LEDGER`, `THERAPIST_ATTENDANCE_HISTORY`, `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY`, `PATIENT_SESSION_HISTORY`, and `DAY_HISTORY_EVENTS` include veryHidden/hidden data-bearing surfaces.
- The broad heuristic marked `_PY_LISTS` as a patient/staff surface because it found role words in the first 30 rows. Treat this as a schema hint only, not proof that the lookup sheet contains patient records.
- A low cached-result count does not prove formula corruption. It means the read-only package scan found fewer cached results than formula cells; actual Excel recalculation/display behavior requires a separate controlled native Excel check.

**Critical limitation:** the supplied console tail contains only the final candidate's worksheet block. The complete four-candidate report has not yet been uploaded/inspected. Therefore no cross-candidate comparison or selection is possible. The four workbooks remain unselected and untouched. No candidate is certified synthetic or approved as the integrated test base.

Next: obtain the full `DIOMIL_H10_STRUCTURE_20261115_165132.txt` file (not another tail excerpt), inspect all four candidate blocks, then compare their worksheet, formula/cache, package and VBA fingerprints against provenance/acceptance evidence. Keep all native Excel write/save operations and operational migrations prohibited at this stage.


## 2026-11-15 — Full four-candidate structural comparison; three distinct VBA projects

The complete uploaded report `DIOMIL_H10_STRUCTURE_20261115_165132.txt` was inspected, superseding the earlier interim note that only the final candidate tail had been received.

### Structural comparison

| Candidate | Whole-file SHA-256 | Size | Sheets | VBA bytes | VBA project SHA-256 | Distinctive sheet |
|---|---|---:|---:|---:|---|---|
| `UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm` | `4df997b6a5e5b0b56b37127ee00dd8710cc26c7c118f1742d58c9969bcbe7afa` | 629,797 | 26 | 920,576 | `8613575ee50826394f336e4f2fd864579f087a2b1e2da4fc9883582e54a1e800` | `RECURRING_ASSIGNMENT_IDS` A1:C180, 540 populated cells |
| `UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm` | `fa9765da7079137bcb1ff935c3d08bc84b409c8335f9cd719f3583b89e4fc67d` | 622,070 | 26 | 921,088 | `7742d872c953f2b026f66097f5a61df84ef18ac151b523e88e901f10bc5a0193` | `DIOMIL_NATIVE_PROBE` A1 only |
| `UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm` | `1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13` | 731,415 | 25 | 1,305,600 | `285b7f4aa20a195535d300928f26b2884255797c1817d4f2a7e32ac82b4033b9` | no extra sheet |
| `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` | **identical to preceding row** | 731,415 | 25 | 1,305,600 | **identical to preceding row** | byte-identical duplicate |

There are therefore **three distinct workbook/VBA fingerprints**, not four. The third and fourth candidates have identical whole-file SHA-256, not just identical VBA projects. They are byte-for-byte duplicates.

All four report 5 defined names, 11 drawings, 2 charts and 2 VML drawings; no external links, connections, query tables, pivot tables/caches, slicers, ActiveX, embeddings or custom UI were reported. The three structures share the same core 25-sheet application layout. Candidate 1 adds the recurring-assignment identity ledger; candidate 2 instead adds a one-cell native diagnostic probe.

Common formula/cache metrics:
- `PATIENT_PLANNER`: 2,745 formulas / 465 cached results.
- `MASTER_SCHEDULE`: 1,021 / 549.
- `STATISTICS`: 43 / 43.
- `PATIENTS`: 549 / 93.
- `OUTPATIENT_SCHEDULE` is header-only; `PATIENT_HISTORY_REPORT` is empty.
- `PATIENT_SESSION_HISTORY` has 523 rows; `THERAPIST_DAY_HISTORY` and `THERAPIST_ATTENDANCE_HISTORY` each have 151 rows; `FINALIZED_DAY_SNAPSHOTS` and `DAY_HISTORY_EVENTS` each have 26 rows.

Low cached-result counts are not by themselves proof of formula defects; Excel recalculation/display behavior requires a separate controlled native check. `_PY_LISTS` is a heuristic role-hint false positive unless separately shown to contain patient/staff rows.

The uploaded report ends with `ALL CANDIDATE STRUCTURE AUDITS COMPLETE`, Python exit code 0, `SOURCE HASH UNCHANGED: YES`, and `NO WORKBOOK WAS MODIFIED`. The initial openpyxl Data Validation warning was non-fatal and did not prevent the completed audit.

### Acceptance and selection boundary

The 2026-11-12 authoritative handoff records:
- CC as the trusted synthetic source and read-only native probe A as frozen;
- synthetic native Save/reopen A as frozen;
- recurring AssignmentID migration A as frozen, with 179 immutable IDs independently verified.

These are evidence/reference artifacts, not writable targets. The recurring-ID workbook has the extra identity sheet and is the best reference for that bounded migration proof; the native-write workbook is the reference for its separate synthetic Save/reopen probe; the CC source/probe pair is byte-identical. None is approved as the full synthetic integration workbook. The whole-workbook synthetic fixture remains NOT CERTIFIED because all patient/staff-bearing surfaces must be proven synthetic and consistently keyed.

### New read-only VBA inventory tooling

Added:
- `Python/tools/audit_vba_project_inventory.py`
- `Tests/test_audit_vba_project_inventory.py`
- `Documentation/H10_VBA_PROJECT_INVENTORY_READONLY_RUNBOOK_20261115.md`

The new native Excel inspector disables macros/events, sets calculation to manual, opens XLSM read-only, lists VBA component names/types/code-line counts and UserForm control names/types, prints no code/cell values/captions, deduplicates byte-identical files and verifies source hashes before/after. It does not save or edit workbooks. VBA project access failures or locked projects are hard stops; do not change Trust Center/AccessVBOM settings to force access.

**Test status:** the last full source regression (1,813 passed, 1 skipped, 26 warnings, exit 0) predates the new VBA inventory tool. The new focused tests and full regression after these commits are NOT YET RUN. Do not report them green until operator evidence is received.

### Next operator gate

Follow `Documentation/H10_VBA_PROJECT_INVENTORY_READONLY_RUNBOOK_20261115.md`: pull latest branch, run the two focused auditor test files, then full source regression, then the read-only native VBA inventory for the four manifest paths. The duplicate CC file will be skipped automatically. If Trust Center access is unavailable, preserve the stop and upload the TXT; do not change security settings. No native write, Save, migration, candidate selection or whole-workbook certification is authorized.


### VBA inventory implementation hardening after initial documentation

The VBA inventory was further hardened before operator execution:
- `audit_one` now verifies source hash/size after every attempted native inspection, including controlled failures; closes without saving; and sets Excel calculation to manual before opening to avoid open-time formula recalculation.
- Duplicate handling now preserves the original audit failure status rather than treating a duplicate of a failed candidate as green.
- An unavailable UserForm control type marks the inventory incomplete (exit 3).
- The runbook's focused pytest invocation now captures stdout/stderr to its own TXT under a temporary `$ErrorActionPreference = "Continue"`, then checks the actual Python exit code.

Relevant source/test commits: `4b24f64c` (inventory tool), `19b730c6` (tests), `19df00b2` (failure/hash/calculation hardening), `07f79294` (duplicate failure preservation), `7a1c01d6` (duplicate-status test), `c0b3f7c0` (UserForm type fail-closed). Runbook updates: `0613de1e`, `6ff8e043`, `21319114`.

Current source/test status remains **NOT RUN after these commits**. The prior 1,813-pass full regression predates this tool and must not be reused as its acceptance evidence. Latest code commit for the VBA inventory: `c0b3f7c0`; all later commits in this sequence update documentation only.


## 2026-11-15 — Project-wide delivery register

The consolidated, evidence-based project delivery report is `Documentation/PROJECT_DELIVERY_MASTER_REPORT_20261115.md`. It brings together the H2–H10 delivery evidence, Recovery track, outpatient/registration/recurring-change workstreams, real-source admission, afternoon scope, protection gates and remaining release blockers. It does not replace specialist proofs or change any frozen artifact. Use `Documentation/CURRENT_HANDOFF.md` for the exact next operator action.


## 2026-11-15 20:31 — H10 call-edge gate accepted; event/validator ownership narrowed

Primary operator artifact: `DIOMIL_H10_CALL_EDGE_BUNDLE_20261115_203111.txt`.

**Observed Windows results:**
- Focused tests: **38 passed in 1.94s**.
- Full repository regression: **1,854 passed, 1 skipped, 26 warnings in 31.08s**; no failures.
- The one skip remains `Tests/test_afternoon_sheet_preview.py:68`, because symlinks are unavailable in the Windows environment.
- Native read-only VBA write-surface audit completed for all three distinct H10 candidates. Each retained `ReadOnly=True`, macros/events disabled, `SOURCE HASH UNCHANGED: YES`, and `NO WORKBOOK WAS MODIFIED`.
- Each candidate still reports 1,569 static operation-pattern matches. Cross-candidate verifier: **GREEN**, three distinct workbooks, 44 nonempty component metadata records, 12 procedure-profile hashes and 9 candidate call-edge references. The pattern count is not a runtime write count.

**Call-edge detail (same in all three candidates):**
1. Five `ThisWorkbook` event calls: Open/Activate → `EnterApplicationShell`; Deactivate/BeforeClose → `ExitApplicationShell`; SheetActivate → `KeepApplicationShell`. These five calls were already matched at exact body-hash level to tracked `application_shell_vba.py::_EVENT_SPECS`.
2. `PATIENT_PLANNER.Worksheet_Change` has three branches, each referencing `CheckSingleAssignment` (native lines 16, 19, 22).
3. `modSchedulingConflict.ConfirmTherapistDoubleBooking` references `ConfigureConflict`.

The native component `Module11` declares `CheckSingleAssignment`, plus `BuildBusyMapsForCheck`, `GetPresentTherapists`, `GetTherapistLoad` and `GetSlotList`. Its 30 detected operation rows are worksheet references and cell-value/formula **accesses** only; no direct assignment or range/structure mutation candidate was detected in this component. Its literal worksheet references are `PATIENT_PLANNER`, `THERAPIST_ATTENDANCE` and `SETTINGS`. This supports classifying it as a read-side single-assignment/availability validation subsystem at the static-pattern level; it is not a complete control-flow proof.

`Module6` exposes `NormalizeTimeMS` and `NormalizePresenceMS`, with identical procedure fingerprints in the three candidates. No candidate VBA call-edge to either helper appeared in this report. This does **not** rule out formula/UDF references, external COM calls, or other invocation mechanisms not represented as a VBA identifier call.

**Next read-only refinement committed after this accepted bundle:**
- The auditor now searches for candidate references to *all* declared VBA procedures in every component with matched operation rows, rather than limiting those components to the two Module6 normalizers. This should reveal Module11's own helper call graph and other cross-module dependencies.
- The structural OOXML audit now counts worksheet-formula references to `NormalizeTimeMS` and `NormalizePresenceMS` without printing formula text, values, identities or cached values. Excel string literals are stripped before matching to avoid false positives.
- The one-command Windows gate now includes this formula scan on the same three frozen candidates and the focused regression for its new privacy-safe parser.

**Current status:** the just-reported 38/1,854 GREEN results validate the preceding revision, not the expanded call-graph/UDF-formula changes committed afterwards. Those newest changes require a fresh focused/full/native read-only gate. No frozen workbook has been modified or selected as a writable integration base.

Next product phase remains the whole-workbook demo. The real-patient source/review workbook and separate eligible pilot XLSM remain a distinct, operator-held admission gate. Do not merge real patient data into a synthetic fixture or treat the H10 candidates as writable targets.


## 2026-11-15 20:40 — expanded all-procedure call graph and formula-only UDF scan prepared

Following the accepted 20:31 bundle, two further **read-only** analysis improvements were committed:

1. `Python/tools/audit_vba_write_surfaces.py` now emits source-redacted `POSSIBLE_CALL_EDGE` candidates against **all declared VBA procedures** from every component with matched operation rows, not only the two Module6 normalizers. This is intended to expose Module11 helper calls and other cross-module edges. Edges are identifier references after quoted strings/comments are removed; they are not runtime execution proof. The synthetic tests now assert cross-procedure discovery and that the previous normalizer-only filter is absent.
2. `Python/tools/audit_integrated_workbook_structure.py` now counts formula references to `NormalizeTimeMS` and `NormalizePresenceMS` in the formula view, per sheet and per workbook, without printing formulas, values, cached results or identities. Quoted Excel string literals are removed before matching. Tests cover actual function calls and false-positive text constants.

`Scripts/run_h10_vba_call_edge_gate.ps1` was updated to include the new focused test file and run the read-only structure/UDF formula scan over the same three frozen H10 candidates. It requires 3/3 completed audits, unchanged source hashes and no-modification attestations, then adds the structure/UDF log to the single output bundle.

**Validation boundary:** these new commits are after the accepted 38-focused / 1,854-full result and have not yet been run by the operator. They must not inherit that earlier GREEN. The next local action is one execution of the updated PowerShell gate with Excel closed. The new script still does not save or modify any XLSM.

After this proof, proceed to the full synthetic integration workbook design/build. The real-source file named `DIOMIL_ΕΛΕΓΜΕΝΗ_ΜΕΤΑΦΟΡΑ_ΠΗΓΗΣ_ΠΡΟΣ_MASTER_20261112.xlsx` was not found among the currently available Project attachments; do not guess its local location or commit it. When the real-data admission milestone is reached, ask the operator to supply the reviewed source privately and follow `REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md`. The A:T column mapping already supplied in the conversation is retained as the operator's source-schema reference; the reclining-treatment replacement rule still requires the operator's fresh clinical instructions at the point of implementation.


## 2026-11-15 20:50 — accepted expanded H10 gate and demo source boundary

The updated gate is operator-accepted: 45 focused tests passed; full regression 1,856 passed, 1 skipped, 26 warnings; all three distinct H10 candidates completed read-only VBA and structure audits, with source hashes unchanged and no workbook modified. The single skip is the Windows symlink limitation. The evidence bundle is `DIOMIL_H10_CALL_EDGE_BUNDLE_20261115_205003.txt` (private operator evidence; not a repository artifact).

The real demo target is the operator-reviewed patient/treatment programme. The private source workbook and identity-review workbook are available in the conversation context; keep them private. The core source mapping is A:T, with additional afternoon/pool fields in U/V. Identity suggestions do not constitute approvals. No real-patient admission or native integrated demo is accepted. The next testable step is a synthetic A:V source-shaped fixture and strict parser/projection roundtrip against the existing trial admission/import planning modules.


## 2026-11-15 — A:V real-source demo schema design

Added `Documentation/TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md` as the current source-to-demo design contract. It records the operator's A:T mapping, the additional U/V afternoon/pool fields, the boundaries of the existing read-only admission/import-plan/weekly-projection modules, and the required architecture for source alignment, patient registry projection, full recurring schedule projection, synthetic end-to-end tests and later native Excel acceptance.

Important unresolved rules are explicit rather than guessed: blank weekday semantics; missing A context handling; U/V operational scope; and fresh operator instructions for reclining replacement behavior at implementation time. The private identity review remains unapproved; no real-patient import or native integrated demo is authorized. Keep source workbooks and all patient-level facts outside GitHub.

 
## 2026-11-15 21:15 — source-alignment audit and cross-platform CI accepted

Added:
- `Python/rehab_excel/trial_source_alignment.py`: read-only comparison of original A:V programme rows, identity review rows and weekly source-cell references. It emits aggregate counts/safe blocker codes only and verifies both input hashes remain unchanged.
- `Python/tools/audit_trial_source_alignment.py`: local private-input CLI; never prints patient names, IDs, room/doctor values, raw source cells or local paths.
- `Tests/test_trial_source_alignment.py`: synthetic tests for exact provenance, identity/context mismatches, wrong source rows/cells, unsupported specialties, missing source times, reclining's time-only source reference, unchanged files and CLI privacy.
- `Documentation/TRIAL_SOURCE_SCHEMA_AND_IMPORT_DESIGN_20261115.md`: A:V field contract and complete source-to-demo architecture.

**GitHub Actions run 43751221120 — GREEN**, commit `94618fa6402c3cfbf028dbcc32b5f694377df754`:
- Focused synthetic safety: **200 passed** on Ubuntu and Windows.
- Full regression: **1,869 passed** on Ubuntu (15 warnings) and Windows (14 warnings).
- No failed or skipped tests in the reported test summaries.

A separate local/private cross-check of the currently uploaded source and review pair found matching patient source-row/name/context provenance and matching weekly source-cell references for the staged review rows. Some A-column ward/doctor contexts are blank and remain blank; no forward-fill was inferred. This check does not approve identities or prove the semantic correctness of proposed weekdays/times/providers.

**Still blocked:** every PatientID confirmation remains pending; suggestions are not approvals; source notes and overlap holds remain unresolved. No real-patient import, native writer or integrated real-data XLSM has been run or accepted. The new alignment CLI is read-only.

**Next engineering slice:** bind the original A:V source hash into the private import plan; build a complete patient-registry plus weekly-treatment projection and synthetic roundtrip; retain U/V as explicitly unmapped until their operational contract is decided. Before interpreting blank weekday cells, obtain the operator's rule. Do not enable import from the existing proposals.

 
## 2026-11-15 — operator decision: weekday-only schedule and blank day cells

Operator clarification: **there are no Saturday/Sunday treatments; the operating treatment week is Monday through Friday (ΔΕ–ΠΑ).** For source import interpretation, when a treatment's corresponding time/value is present but its weekday cell is blank, interpret that treatment as recurring Monday–Friday. Never generate Saturday or Sunday sessions. A missing treatment time/value means no session is created; do not create a weekday slot from a blank day cell alone.

Implementation must keep this as an explicit source-import rule (pass `implicit_daily=True` only for the relevant source mapping when a treatment value exists and its day cell is blank), rather than changing the general parser default for unrelated callers. Add synthetic regression coverage for blank-day + present-time => exactly ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ and zero weekend entries; blank-time + blank-day => zero entries. Preserve explicitly supplied weekday patterns unchanged.


## 2026-11-15 — DEMO priority: source-bound patient registry + weekly programme (engineering checkpoint)

**Operator priority:** build ONE openable integrated practice DEMO for admission/discharge, patient/therapist absences and replacements, new therapists/students, restoring changes and day operations. Statistics, H9 attendance charts, 15-day patient progress and management analytics are subsequent/parallel, **not blockers for the first practice DEMO**. Do not start over with H2–H10 native proofs or repeat previously accepted gates without a code change or new evidence.

**Code delivered on `work/afternoon-medical-scope-20261115` after the already accepted 1,869-pass source-alignment CI:**

- `Python/rehab_excel/trial_source_weekday_projection.py`: pure, source-cell-preserving A:V row expansion. Populated clock + blank days maps to exactly ΔΕ–ΠΑ; blank clock creates no slot; explicit split weekdays/times retained; no weekends. The import-specific `implicit_daily=True` is never made a global parser default. Unsafe/incomplete inputs hold the entire source row, not a partial import. Reclining F and afternoon/pool U/V are **unmapped blockers**, not silently discarded. This is a proposed mapping only.
- `trial_import_plan.py`: optional `programme=` requires original source alignment and binds its SHA256 alongside review and eligible intake. The old two-input plan remains readable but **is not source-bound**; it may not be used by the new registry projection. A source-bound plan requires the programme again for `verify_sources` and blocks any changed/missing source.
- `Python/rehab_excel/trial_patient_registry_projection.py`: read-only `PATIENTS` A:E + optional type/MRN/doctor and separate `SourceContext` in-memory rows, with exact original A/B source row and review confirmation checks. Clinical `PatientType`, infectious yes/no, and any room/doctor/status must be independently provided by an authorized reviewer, never derived from names, section headings, ward/doctor context, or guessed formatting. Blank source context stays blank. Nothing writes patient data to GitHub.
- `Python/rehab_excel/trial_demo_source_projection.py`: combines that registry with existing `PROGRAM_WEEKLY_SLOTS` projection, requiring all weekly PatientIDs to exist in the explicitly classified inpatient registry. Outpatient weekly entries remain blocked under this inpatient-only schema rather than silently omitted.
- Synthetic tests: `Tests/test_trial_source_weekday_projection.py`, `Tests/test_trial_patient_registry_projection.py`, plus original programme SHA/legacy fail-closed cases in `Tests/test_trial_import_plan.py`. Tests use invented identities and do not open real source workbooks.

**Boundary:** This is a read-only, source-level implementation milestone; it is **not yet a native XLSM importer, integrated practice workbook, confirmed real patient migration, certified operational demo or clinical-data approval**. Do not use frozen H10 source/proof workbooks as writable destinations. CI result for each new commit must be confirmed separately; the older 1,869/200 GREEN cannot be inherited by new changes.

**Next development sequence:** independent source-bound projection fixture and exact weekly/registry reconciliation -> eligible disposable native workbook writer and independent Windows Excel readback -> integrate the already existing admission, discharge, replacement, therapist/student registration and in-session UX within the same file -> operator practical scenarios. Actual reviewed source and identity approval stay private; no approval means no real patient import. Keep demo repeatable via a clean, frozen starting copy. Historical resume/test instructions above this appendix remain chronology only where contradicted by this newer checkpoint.


## 2026-11-15 — DEMO A:V source versus admitted weekly parity GREEN

**Priority remains one integrated practice DEMO first**, covering patient admission/discharge, absences, replacements, new therapists/students and daily operations. History, H9/H10 statistics, 15-day patient progress and management reporting are later/parallel; they are not prerequisites for the practice workflow.

**Delivered:** `Python/rehab_excel/trial_demo_source_reconciliation.py` independently reconstructs every proposed weekly entry from the original A:V programme via `trial_source_weekday_projection.preview_source_weekdays`. For each explicitly confirmed patient identity, it compares the complete multiset of (PatientID, treatment, weekday, time, provider, source-cell) against `TrialImportPlan.entries`, rejecting missing/extra/changed sessions. This catches an important limitation of the earlier source-alignment audit: correctly pointing to C4 did not prove that the staged clock/day/provider were correctly copied. Blank weekdays + populated time expand to exactly ΔΕ, ΤΡ, ΤΕ, ΠΕ, ΠΑ. No sessions are created without time. If any F reclining or U/V afternoon/pool source fact is not mapped, or any supported cell is ambiguous, the entire patient row is held rather than imported partly.

**Integrated gate:** `trial_demo_source_projection.project_trial_demo_sources` now checks this exact source parity **before** combining PATIENTS and PROGRAM_WEEKLY_SLOTS rows. Synthetic acceptance tests added to `Tests/test_trial_patient_registry_projection.py`: source blank-day five-weekday parity, changed time under correct source-cell reference, blocked unmapped afternoon, and existing combined source/registry/weekly synthetic roundtrip. No real clinical data, native write or VBA modification was made.

**Evidence:** GitHub Actions `43758856408`, code commit `ba001d35459732724ca2c9be7f51b3864720d6d7`, completed **all four Windows/Linux synthetic and full regression CI jobs successfully**. An intermediate code-only commit before the synthetic fixture was updated had a red CI run; that intermediate result is superseded by the green `ba001d35` run. Do not rerun this accepted source gate merely for reassurance. It is not native Excel acceptance.

**Known boundary / precise next implementation:** only source-bound read-only projection and reconciliation exist, not the requested runnable XLSM. A native writer requires a fresh eligible **empty operational demo shell** and independent Excel Save/Close/Reopen, ensuring the old identities, frozen H10 artifacts and history remain untouched. Since `PATIENT_PLANNER`/MASTER is visual single-slot oriented while `PROGRAM_WEEKLY_SLOTS` supports multiple distinct day/time placements, the writer MUST expose complex weekly rows without compressing them to an incorrect single time. Confirmed private PatientIDs, clinical classification and holds/shared-provider approvals remain mandatory before importing real patient records. Target one independent practice workbook, not another analytics prototype.


## 2026-11-15 — Private clean-shell native DEMO seeding bridge

**Immediate delivery goal remains a single runnable practice workbook** for patient admission/discharge, absences, therapist replacements, new therapist/student registration and daily operations. Statistics/15-day progress remain subsequent/parallel work. Historical H2–H10 accepted proof artifacts remain frozen and are NOT the write destination.

**Implementation added**:
- `Python/tools/preflight_trial_excel_sources.py`: opt-in `allow_empty_demo_shell=True` permits ONLY an otherwise eligible macro-enabled intake containing **zero patients, zero protected history records** and exactly the `EMPTY_OR_MISSING_PATIENT_REGISTRY` readiness issue. Default preflight behavior stays unchanged. Any real patient registry, external workbook links, missing required operational sheet, malformed history or other issue continues to block.
- `Python/rehab_excel/trial_import_plan.py`: `empty_demo_shell=True` must be separately proved by preflight; `TrialImportPlan.empty_demo_shell` is not inferred and default remains False. An eligible real-source DEMO still requires independently confirmed IDs, zero review holds, zero overlaps and exact shared-provider approval.
- `Python/rehab_excel/trial_demo_native_seed.py`: guarded `seed_trial_demo_workbook` creates a unique **private output only outside a Git checkout**, from an eligible empty shell, through one isolated temp XLSM. It refuses existing output files, old patient/planner/weekly/other operational records, students, historical rows or source-hash changes. Real XLSM modifications use Excel COM, not openpyxl save. It writes PATIENTS, row-bound PATIENT_PLANNER identity formulas, granular PROGRAM_WEEKLY_SLOTS, and a full non-compressed DEMO_WEEKLY_VIEW; native Excel saves and closes before independent readback. The output is exclusively published only after verification, with no overwrite. It verifies the data read back by the actual `read_base_schedule` daily-scheduling source reader, protecting replacement-engine input against phantom/missing weekly slots. Optional mock backends are never considered proof of native Excel acceptance.
- Tests: `Tests/test_trial_demo_native_seed.py`, extensions in `Tests/test_preflight_trial_excel_sources.py` and `Tests/test_trial_import_plan.py`. These use synthetic fake VBA containers solely for structural rejects and no clinical identities.

**Explicitly NOT delivered/accepted:** no populated private real-source XLSM, no native Windows COM Save/reopen trial on the user's machine, no full UI buttons/workflow acceptance, and no patient-ID authorization. The clean shell has not been created/approved for this specific import. The legacy MASTER/PATIENT_PLANNER visual model is single-slot oriented and cannot correctly compress all multi-day/time treatment cases; DEMO_WEEKLY_VIEW preserves the complete schedule but is **not** a substitute for integrating all complex treatments into the familiar operational UI. Do not label this partial native bridge a finished practice application.

**Next true engineering gate** (not repetition of old H2–H10 tests): prepare a separately eligible EMPTY operational demo shell, bind original A:V/review/source-hashes and approved private clinical facts, perform a fresh Windows native seed on a never-existing private path, verify the workbook within Microsoft Excel and exercise admission/discharge, therapist absence and replacement, therapists/students and daily restore. Do not rerun already accepted 1,869+ synthetic history tests merely for reassurance; only new changes require scoped regression. See `Documentation/PRACTICE_DEMO_NATIVE_SEED_RUNBOOK_20261115.md`.


## 2026-11-15 — Production reader readback added to native DEMO seed

The native seed's independent post-save verification now also calls the production `read_patient_registry(stage)` parser and compares every imported PatientID, display name, room, infectious flag, status, inpatient/outpatient classification, hospital MRN and responsible doctor with the approved in-memory registry projection. This sits alongside the existing `read_base_schedule(stage)` comparison for every granular recurring treatment/day/time/provider. This closes a readback gap: matching worksheet cells alone did not prove that both production source readers could consume the native result.

Code commit: `702db942b5e9416344438f744efa7a6889dcc5e7`. This commit was made after the previously green `2b1c85bd` workflow; **do not claim the new reader-check commit has CI evidence until a run for this exact SHA completes**. No private source workbook, native Excel run, or operational UI acceptance was performed.


## 2026-11-16 — DEMO patient exit button, preview, explicit confirmation and native seed integration

**Implemented (source + synthetic CI, not yet native operator acceptance):**
- `Python/rehab_excel/trial_demo_patient_exit_vba.py` installs a separately named **Εξιτήριο / Λήξη** button on the MASTER, preserving existing MASTER toolbar and `Πλήρες εβδομαδιαίο` button. It installs `modDemoPatientExit` with `DemoPatientExitAction` only on the isolated DEMO candidate. The macro refuses non-`_DEMO_` filenames or read-only sheets, requires exactly one authorized PatientID, validates the registered inpatient/outpatient type and the existing `PATIENT_EXITS` schema/IDs, requests an exact YYYY-MM-DD *last active day* and optional reason, previews the affected recurring entry count, and requires a default-**No** explicit confirmation. Only then does it append one patient-exit event, save in Excel and recheck the in-session sheet. If writing/save/validation fails, it attempts to clear the newly appended event and save the reversal; it must still be independently reviewed after any failure. Existing PATIENTS, PATIENT_PLANNER, PROGRAM_WEEKLY_SLOTS, outpatient source and finalized history are NEVER erased by this action. It does not silently equate an exit with a daily absence.
- `Python/rehab_excel/trial_demo_native_seed.py` now installs this new UI module and verifies the macro and button action via read-only Excel COM before publishing a **new isolated** candidate. The original source and frozen H10 proof files remain untouched.
- `Python/rehab_excel/patient_exit_source.py` now also offers `verify_saved_patient_exit`, an independently reopening **read-only** verification helper: confirms exact recorded kind/date and original recurring rows still exist, but those rows disappear for that patient after the last active day. This **is not automatically invoked by the live VBA macro**, whose immediate validation is in-session only.
- Synthetic tests in `Tests/test_trial_demo_patient_exit_vba.py` and `Tests/test_patient_exit_source.py` cover button wiring, repeatable module installation, confirmation guard, rollback markers, original recurrence preservation and wrong date/type rejection.

**Next operator/native gates:** seed a private fresh eligible XLSM using approved identities and actual Windows Excel; click exit and cancel, then confirm on a separate disposable case; Save, Close, Reopen and independently read back `PATIENT_EXITS` and subsequent DAILY_INPUT for the right effective date. The present MASTER census/visuals are **not automatically refreshed** by this new VBA action, and the macro explicitly warns the operator. Do not label the exit workflow end-to-end complete until those checks, active MASTER census and menu navigation are proven. Python/Windows CI is source regression evidence, not native Excel acceptance. No patient data must be committed to GitHub.


## 2026-11-16 — DEMO effective MASTER and live patient census bridge

**Implemented on `work/afternoon-medical-scope-20261115`**:
- `Python/rehab_excel/master_projection.py`: `build_master_projection(..., target_date=date)` optionally reads the validated `PATIENT_EXITS` ledger and excludes an inpatient from the projected **active MASTER census only after** their inclusive last-active date. Without `target_date`, legacy MASTER behaviour is unchanged. Patient registry and recurring schedules/history remain immutable. The preview fails closed on invalid dates or ledger identities.
- `Python/rehab_excel/master_projection_writer.py`: optional `target_date=` on `refresh_master_projection_in_place` with standard independent readback. This creates an explicit opt-in native MASTER regeneration path, **not yet invoked from the in-session DEMO discharge action**; do not claim automatic MASTER census refresh.
- `Python/rehab_excel/trial_demo_active_census.py`: a distinct, formula-backed `DEMO_ACTIVE_CENSUS` status view linked directly to PATIENTS and PATIENT_EXITS, with its own **Ενεργοί ασθενείς** button in MASTER and **← MASTER** back action. Date cell `H2` defaults to Excel `TODAY()` and can be changed to inspect past/future dates. Every patient remains visible with status `Ενεργός` or `Εξιτήριο / Λήξη` as of H2. Last active day is inclusive. PatientID lookup coerces IDs to text so new numeric registrations can match text-stored exit events. Formula column references follow the actual PatientType column. The workbook's `RefreshDemoActiveCensus` VBA extends source-linked formulas when new PATIENTS rows are registered; a blank row-2 template exists even with zero patients.
- `trial_demo_native_seed.py` installs this view and checks its formulas, button actions and VBA module in read-only native Excel COM, in the same isolated staged candidate as the existing weekly and exit buttons. `trial_demo_patient_exit_vba.py` invokes `RefreshDemoActiveCensus` before saving a confirmed exit, so a new patient's exit is reflected without deleting anything.
- Tests: `Tests/test_trial_demo_active_census.py`, additions to `Tests/test_master_projection.py` and `Tests/test_master_projection_writer.py`, covering inclusive dates, valid/invalid requests, immutable registry, dynamic formula IDs and future admission rows. All tests use synthetic identities only.

**Important status boundary:** Python CI must be checked for the final commit. The operator-facing in-session discharge action STILL does **not** automatically rebuild the canonical MASTER display (it now refreshes the separate DEMO_ACTIVE_CENSUS). The candidate has **not** been created from private approved patient workbooks or tested by the operator in native Windows Excel; no claim of operational acceptance. Do not change canonical MASTER automatically without native Save/Close/Reopen/rollback evidence. Complete admissions/replacements/staff/student lifecycle acceptance remains pending. No patient-identifying source or workbooks go to GitHub.

**Next priority:** safe in-session native MASTER census refresh of the chosen target date (without changing historical rows), followed by one executable private DEMO and full operator flows. Previously green H2-H10 proof checkpoints are frozen; do not rerun unrelated history/stats tests. `Documentation/PRACTICE_DEMO_NATIVE_SEED_RUNBOOK_20261115.md` is the operator safety runbook.


## 2026-11-16 — F/U/V read-only gate, 1,991 tests GREEN

**Concrete new gate, not a synthetic replacement for the user's real-data DEMO:** `Python/rehab_excel/trial_special_fields_audit.py`, `Python/tools/audit_real_manual_specials.py`, `Tests/test_trial_special_fields_audit.py` provide the first **single-source-only** private read-only inspection of actual manual A:V columns **F** (reclining, known 30-minute ΦΘ-related operation), **U** (legacy afternoon, no auto-import) and **V** (morning pool). No review/clinical approval/clean XLSM required for this *diagnostic* first command. It checks original source sheet/header, SHA-256 before/after, cached and raw Excel formula values (never omit formula-only cells), missing-patient source rows, time parseability, all counts and no PII emitted. The tool **never creates sessions, guesses F/V weekdays or imports old U**. It returns 0 if F/U/V truly empty (not whole source approval), 3 for pending real special records, 2 on safety stop. Exact local command, interpretations and privacy gates: `Documentation/REAL_MANUAL_SPECIAL_FIELDS_AUDIT_20261116.md`.

**Verified code evidence:** `95157af7dbb32c6cb8aca210c503ffbb7f63be17`, GitHub Actions `43804737124`, **1,991 passing pytest in both Windows and Linux**, all four full/synthetic CI jobs successful. This is not proof that the real private manual has been scanned or a working native Excel was built. **Real-data DEMO remains NOT DELIVERED.** F/V day/provider/duration-aware scheduling approval, U source scope and independently reviewed outpatient programme remain blockers. Next meaningful step is run the read-only audit against the operator-approved private CURRENT manual snapshot, capture only totals, finish explicit F/V mapping and standalone outpatient source, then eligible private native seed and full side-by-side weekday/operator acceptance. Frozen H2–H10 proof artifacts remain untouched.


## 2026-11-16 — Mandatory F/U/V source gate before real DEMO seed

`Python/tools/prepare_real_data_practice_demo.py` now calls `inspect_real_manual_special_fields(programme)` **before reading approved clinical values, building source plans or invoking native Excel**. Any populated F reclining, U legacy afternoon or V morning pool field (including Excel formula-only content with empty cache), or an orphaned source row, gives a privacy-safe **SAFETY STOP**. This ensures that even a stale review workbook cannot silently hide real original-source special therapies. F/V clock parseability is diagnostic only: **days/duration-aware scheduling and authoritative source treatment review are still required**. U remains expressly non-importable from the historical column; separately entered authorized afternoon appointments belong to their own programme. The standalone command `python Python\tools\audit_real_manual_specials.py --programme "<PRIVATE_CURRENT_MANUAL>.xlsx"` reports only aggregate counts and returns 3 for blocked real specials, 0 if no F/U/V cells are populated, 2 on structural safety failure. No original workbook or native XLSM was changed in the GitHub development environment.

**Tests:** `Tests/test_prepare_real_data_practice_demo.py` adds explicit refusal of preflight and native build before clinical approval/Excel on pending F/U/V/orphan formula cells, alongside `Tests/test_trial_special_fields_audit.py`. Verified integrated hard-stop code checkpoint `a9d2f0afd8b119b69a562475525d30d35d7cd8a3`: **1,993 tests passed on Windows and Linux, 4/4 CI jobs successful** in GitHub Actions run `43805000757`. This is source/synthetic regression evidence, not a real-manual scan or native Excel acceptance. Real private source inspection, source review, native Excel candidate and user acceptance remain **PENDING**. The new hard-stop is deliberate and must NOT be bypassed with invented patients or source omissions.

## 2026-11-16 — Existing MASTER content and formatting reused in real-data DEMO

**User clarification confirmed in source, NOT a new MASTER design:** `Python/rehab_excel/layout_contract.py` defines MASTER headers, treatment locations (D ΦΘ, E ρομποτικό, F πισίνα, G ανακλινόμενο, H εργο, I λογο, J ΕΦΑ, K ψυχολόγοι, L απογευματινό, M κατάσταση). `patient_program_master_render_vba.py` already implements `RefreshMasterPatientProgram`, `BuildTreatmentDisplay`: the actual treatment cell content is **clock (top line), weekdays (next line), optional approved provider (third line)**, wrapped and centered; `master_visual_style.py` owns existing colors, dimensions, header/identity styling, clinic banners. `PATIENT_PLANNER` has each matching `<treatment>_Ώρα`, `<treatment>_Ημέρες` and optional provider. These contracts are preserved rather than re-created.

**Integration defect discovered/targeted:** `master_projection_writer.py` must first sort admitted patients; its original one-cell formula templates, when used by the trial seeder, do NOT guarantee the already established multiline treatment display. New `Python/rehab_excel/trial_demo_master_display_bridge.py` reopens ONLY the disposable private staged `NATIVE_DEMO_STAGE.xlsm`, **requires** the already-installed `PatientProgramMasterRender.RefreshMasterPatientProgram` VBA module, calls it for each imported inpatient after the existing MASTER projection + visual style, and saves. It **does not install a duplicate formatter, change cell layout, modify the operational original or touch H2-H10 proofs**. Independent read-only `verify_private_demo_master_display` checks exact time, weekdays, optional provider and Excel WrapText for each losslessly represented (patient, treatment) in the initial trial; mismatches or a missing established VBA macro **fail closed** before publishing. The existing `DEMO_WEEKLY_VIEW` still shows all weekday/time variants where the legacy single-cell representation cannot show them losslessly. Import of actual source F/V/U special treatment fields is a SEPARATE unresolved input-admission contract; **none of the already supported treatment categories is prohibited**.

**Proof scope:** `Tests/test_trial_demo_master_display_bridge.py` checks original VBA column placements, full content/format order, sample persisted readback, deliberate tampering, missing/stale macro and seeder call order. CI for this latest commit is tracked separately. Native Windows/COM real-data candidate has **not** yet been executed. Required next actions are exact real manual source F/V/outpatient approval, private native seed and operator comparison, not another MASTER redesign.


## 2026-11-16 — Real outpatient source audit reuses existing application schema

**Current verified architecture:** Normal patient registration/program editor already writes `PATIENT_PLANNER` and uses `PatientProgramMasterRender.RefreshMasterPatientProgram` for the existing MASTER cell content. The new bulk real-data DEMO seed is a *different* initial-population path; it does not imply the old MASTER renderer was broken. Do NOT rebuild established cell formatting, controls or clinical treatment categories.

**Next separate real-data source bridge:** `Python/rehab_excel/trial_private_outpatient_audit.py` and `Python/tools/audit_real_outpatient_source.py` inspect only a PRIVATE, institution-approved current outpatient export in `.xlsx` form. They require the already-supported `PATIENTS` and `OUTPATIENT_SCHEDULE` sheets and their existing exact leading column headers. Reuse `read_patient_registry`, `read_outpatient_schedule` and `parse_day_pattern` to check outpatient identity/type/name consistency, Mon-Fri days, times, provider for ΦΘ, no robotic outpatient, per-weekday duplicate bookings, no formula-only or orphan source rows and SHA-256 read-only preservation. CLI emits only aggregate counts; never names/IDs/paths. `Tests/test_trial_private_outpatient_audit.py` uses only fictional private fixtures. **This is a diagnostic admission precondition, not an outpatient import or proof of completeness**. Full native outbound bridging from this separate approved source into the same private XLSM remains PENDING. Do not infer outpatients from the existing inpatient A:V review or omit them from the final pilot. Full command and schema: `Documentation/REAL_OUTPATIENT_SOURCE_EXISTING_SCHEMA_AUDIT_20261116.md`.

**Real workbook and operator acceptance status remains unchanged:** no current actual source files available in GitHub, no private native Excel produced or verified, outpatient source admission/merged seed and F/U/V approvals still open, 1:1 real manual comparison pending. Previous 2,002 test green code checkpoint `3cfdf61b` with GitHub run `43805735643` establishes the earlier MASTER bulk-display integration, not outpatient native acceptance. Check the new outpatient audit CI separately. Frozen H2–H10 acceptance files stay immutable.


## 2026-11-25 — Current candidate Day Amendment preflight exception

Candidate `UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` retained its recorded baseline SHA-256 after read-only inventory. On this artifact, the UI contract inspector reports the expected day-save button missing from `DAILY_INPUT`; the in-session persistence diagnostic stops earlier because the candidate is missing required VBA component `modDailyInputJournal`. Thus this candidate is not presently eligible for positive Day Amendment execution. No successful mutation, receipt, native save/close/reopen, or rollback has been demonstrated on CC. This does not invalidate the separately frozen H3/H5/H6 native amendment evidence, whose exact scopes and artifacts remain authoritative for those historical tests. Required next work is read-only control/component inventory and reconciliation of installer, inspector and application-contract expectations. Preserve all baselines.


## 2026-11-25 — CC native inventory / UI placement findings

Read-only native inventory of CC completed successfully (66 components; hash unchanged). `modDayAmendmentAction` and `modUnifiedDaySave` are present; `modDailyInputJournal` is absent. The unified operational-day save shape exists on `MASTER_SCHEDULE` and points to `SaveOperationalDay`; no matching day-save/amendment shape exists on `DAILY_INPUT`. The Day Amendment UI inspector therefore fails its DAILY_INPUT button expectation, while the separate unified-save inspector passes. No workbook mutation or transaction occurred. This distinguishes a sheet/control placement contract issue from the separate missing-journal dependency. Next step is source-level reconciliation of installer, control constants and intended workflow, without installing or modifying anything.


## 2026-11-25 — Source-level Amendment UI contract

The source defines the Amendment button as `ekaToolbar_DayAmendment` bound to `StartClosedDayAmendment` on `DAILY_INPUT`. Its installer replaces `modDayAmendmentAction`, creates that button on `DAILY_INPUT`, verifies it and saves. The separate unified-save button on `MASTER_SCHEDULE` does not satisfy this Amendment contract. The Amendment VBA calls `DailyInputJournalApplyEnabled()`; the current CC inventory lacks `modDailyInputJournal`, and the Amendment UI installer does not install it. Do not run the installer on CC until the journal dependency/provisioning contract is resolved. No workbook mutation performed.


## 2026-11-25 — Journal PREPARED-only state and UI installer conflict

The journal native provisioner creates/validates the protected journal sheet only and explicitly installs no operational hooks. The separate `modDailyInputJournal` source is a PREPARED-only bridge; `DailyInputJournalApplyEnabled()` deliberately returns False, and the complete Apply/Restore state machine is not enabled. Journal prepared preflight is read-only; transition proposals remain non-authorizing. The Unified Save installer centralizes saving on `MASTER_SCHEDULE`, routes to CloseDay or Amendment handlers based on snapshot state, and explicitly deletes the DAILY_INPUT Amendment button. This conflicts with the standalone Amendment UI installer that expects the button on DAILY_INPUT. The integrated application UI contract and module provisioning sequence must be reconciled before any installation. Do not modify original CC; no positive operational transaction is yet ready or proven.


## 2026-11-25 — DAILY_INPUT lifecycle evidence review

PREPARED-only journal VBA bridge is present in source; Apply gate deliberately returns False. Pure lifecycle state model defines legal phases but does not persist or orchestrate Excel effects. Evidence policy enumerates native reopen obligations but confirms no independent phase verifiers are registered and does not authorize writes. Therefore operational transaction orchestration, positive Apply/Restore/Amendment, phase receipts and full-workbook rollback remain NOT IMPLEMENTED / NOT PROVEN. Keep Apply gate closed. Next: inventory handlers and persistence boundaries, implement independent verifiers and VBA orchestration, then test rollback/fault injection on disposable synthetic native Excel copies.


## 2026-11-25 — Handler inventory and rollback scope

Existing source handlers identified: Apply (`modDailyInputAction`), automatic replacement (`modGlobalReplacementAction`), manual replacement (`modReplacementOperator`), ledger/replacement Restore (`daily_restore_in_session_vba.py`), Amendment (`modDayAmendmentAction`), and strict full-workbook rollback audit (`daily_input_full_rollback_audit.py`). Operational paths remain locked by `DailyInputJournalApplyEnabled=False`.

Local rollback attempts already exist, so “no rollback code exists” is inaccurate. However, Apply's outer `IncompleteApply` branch retains the operation snapshot and warns about partial saved changes rather than automatically restoring it. Other local copy-back/restore paths explicitly disclaim independent native Save/Close/Reopen proof. The full-package audit is read-only and is not a rollback executor; it also requires separate trusted receipt and native Excel reopen. Integrated durable phase orchestration, recovery and positive native transaction proof remain NOT ESTABLISHED. Keep gate closed.


## 2026-11-25 — Test coverage map: handlers, evidence classes and positive-proof gap

Reviewed the focused test set for Apply, Gate 2, rollback audit/native diagnostic, automatic and manual Replacement, in-session Restore, and Day Amendment. Existing tests cover:
- Source contracts and mutation-guard ordering (`test_daily_input_action_vba`, `test_daily_input_gate2_guard_matrix`, `test_global_replacement_action_vba`, `test_day_amendment_action_vba`, `test_daily_restore_in_session_vba`, `test_replacement_operator_recovery`).
- Synthetic package/journal audit and CLI wiring (`test_daily_input_full_rollback_audit`, `test_verify_daily_input_full_rollback`).
- Disposable native Excel negative/fault-injection diagnostics (`test_diagnose_daily_input_gate2_apply_worker_native`, `test_diagnose_daily_input_gate2_restore_worker_native`, `test_diagnose_daily_input_gate2_day_amendment_worker_native`, `test_diagnose_daily_input_full_rollback_native`).

These are distinct evidence classes. Source tests do not execute VBA; synthetic fixtures do not prove native workbook behavior; negative native tests prove rejection/fault-recovery behavior only. The reviewed set does not establish a positive permitted public-UI Excel transaction with durable receipt, Save/Close/Reopen readback, dependent projection/history verification and independently proven rollback. Gate 2 remains false.

Next: map durable journal state transitions and trusted receipt binding across public entrypoints and mutation boundaries; close the orchestration gaps and verify disposable native fault injection before any positive operational transaction test.

## 2026-11-25 — Durable journal phases and receipt chain audit

Source-confirmed lifecycle: PREPARED → INPUT_AND_LEDGER_COMMITTED → PROJECTIONS_VERIFIED → optional PENDING_OPERATOR_DECISION → REPLACEMENTS_RESOLVED → APPLIED_ARCHIVE_VERIFIED → COMPLETE. Failure path is FAILED → ROLLED_BACK, with verified rollback required. COMPLETE and ROLLED_BACK are terminal.

The append-only journal codec validates OperationID, sequential Revision, Phase, source hash/service date and checksummed canonical payload. It prevents a new operation before terminal verification. Its SHA-256 checksum detects corruption, not authenticity.

Native VBA journal support is PREPARED-only at present. The apply switch returns literal False. The PREPARED bridge saves and independently reads back the PREPARED row but does not execute operational Apply or persist later phases. The pure phase-evidence policy lists native reopen requirements but explicitly has no registered independent verifiers; it cannot produce proof or authorize writes.

Trusted receipt chain is not yet established. The rollback auditor expects a prepared-snapshot hash supplied as expected_snapshot_sha256, while its own documentation requires that digest to be bound by a separately verified durable transaction receipt. The auditor does not authenticate that receipt. Thus a caller-supplied digest cannot stand in for receipt authenticity.

Implementation order: receipt schema and source/operation binding; independent native reopen verifiers for every evidence kind; VBA phase orchestrator; full rollback executor plus independent pre-terminal and terminal proof; disposable synthetic native fault-injection across every write/save/refresh boundary. Only after all are green may a positive public-UI transaction be considered on a disposable demo. Gate remains false.

## 2026-11-25 — Transaction Receipt System design baseline

Design-only contract added at Documentation/DAILY_INPUT_TRANSACTION_RECEIPT_DESIGN_20261125.md. Proposed canonical receipt binds operation/source identity, service date, journal revision/phase, phase-required evidence kind, relative artifact locator and exact SHA-256, pinned verifier identity/version/method, UTC time, previous receipt digest and receipt digest. Native workbook evidence requires independent Excel reopen. Receipt and journal chain must match; incomplete crash-window states fail closed.

SHA-256 alone is not authenticity. No trusted signer/attestation root exists in the reviewed repository; no key is embedded or invented. Structural receipt validation may be developed first, but receipt design is not authorization and Gate 2 remains false. No runtime code, VBA or workbook was changed.

## 2026-11-25 — Structural receipt codec/test addition

Added daily_input_transaction_receipt.py and test_daily_input_transaction_receipt.py. Coverage targets canonical JSON/digest, strict schema, hashes/timestamp, path traversal/symlink containment, operation/phase/evidence binding, verifier registry labels, predecessor chain and exact evidence artifact bytes. Module is non-authorizing: digest is not authenticity, registry labels require trusted build-manifest provenance, and no journal writes/Gate enablement are exposed.

Tests were added but NOT RUN in this session. Required next action is local pytest execution for the new module plus journal, phase-evidence and rollback-audit regression suites. No workbook or VBA changed; Gate 2 remains closed. Trusted signing/attestation root and actual native phase verifiers are still absent.

## 2026-11-25 — Receipt PREPARED evidence-kind follow-up

Initial focused run: 68 passed, 1 skipped, exit code 0, with six PytestUnraisableExceptionWarning instances involving ZipFile deallocation in full rollback audit tests. Follow-up source review corrected PREPARED evidence semantics: new EvidenceKind.PREPARED_SNAPSHOT is required only for revision 0; FAILURE_DIAGNOSTIC is not a valid PREPARED receipt kind. Added one test covering both invalid phase/revision combinations. This follow-up revision is NOT YET TESTED. One skip likely relates to symlink support on Windows; confirm with pytest -rs. Warnings are recorded for investigation, not suppressed.

## 2026-11-25 — Receipt fixture follow-up run and correction

After PREPARED_SNAPSHOT enforcement, the local combined suite reported 5 failed, 64 passed, 1 skipped, with six ZipFile unraisable warnings. Failures were test setup inconsistencies: later-phase receipt tests still generated revision-zero PREPARED receipts, and the missing-predecessor test did not select later-phase evidence. Test fixtures were corrected to represent revision 1 INPUT_AND_LEDGER_COMMITTED / INPUT_AND_LEDGER_REOPEN and a PREPARED predecessor digest. The missing-predecessor test now removes only that digest. Correction is committed but NOT YET RERUN. Skip reason: symlinks unavailable on Windows. ZipFile warnings remain open investigation; do not silence them.

## 2026-11-25 — Receipt suite green; rollback ZipFile warning follow-up pending

Focused combined suite: 69 passed, 1 skipped, 6 warnings in 5.22s. Symlink escape test skipped because Windows symlink creation is unavailable. The warnings are PytestUnraisableExceptionWarning from ZipFile.__del__ in full rollback audit cases.

A source-level resource-lifecycle correction was committed: _read_journal() now calls openpyxl load_workbook with keep_vba=False because it is read-only and VBA package preservation is unnecessary there; package VBA presence is verified separately. Commit 4a7c168dfecb2eb30de79ae0dc87c3d3c4b763ec. This change is NOT YET TESTED. Re-run focused tests with warnings visible; do not filter or suppress them.

## 2026-11-25 — ZipFile warnings narrowed to synthetic fixture keep_vba setting

Latest local result: 69 passed, 1 skipped, 2 warnings in 4.37s. Warnings remain only for formula/style variants of the synthetic business-cell mutation test. The test fixture is a non-macro openpyxl workbook with .xlsm extension; two test loads used keep_vba=True unnecessarily. Both were changed to keep_vba=False in test code (commit 30f278bf1fe6d624fe7774ac51c3a6f3466a6378). This is a test-only cleanup, not a production workbook change. Rerun pending; warnings remain open until confirmed absent. Symlink skip remains Windows environment limitation.

## 2026-11-25 — Phase-verifier contract/report tests added — execution pending

New files: Python/rehab_excel/daily_input_phase_verifier_contract.py and Tests/test_daily_input_phase_verifier_contract.py. Synthetic coverage includes all ten EvidenceKind registry entries, immutable registry, canonical report/digest, tamper and schema rejection, caller Boolean rejection, unsafe locators, transition-derived evidence obligations, identity and scope mismatch, rejection of non-VERIFIED outcomes, explicit fail-closed behavior while PREPARED remains diagnostic-only and the other actual verifiers remain NOT_IMPLEMENTED, and rollback package/PREPARED snapshot requirements.

These tests have NOT been executed locally. Latest reported focused baseline before the two test-only keep_vba fixture changes: 69 passed, 1 skipped, 2 warnings. The fixture correction is unverified until rerun. Do not claim phase verifiers operational; this addition is contract metadata/report validation only.

## 2026-11-25 — Confirmed Windows focused baseline after phase-verifier contract/report addition

User ran the focused suite and reported:

```
110 passed, 1 skipped in 4.79s
SKIPPED [1] Tests\\test_daily_input_transaction_receipt.py:208: Symlinks unavailable on this platform
```

Confirmed green: 110 passed, zero failed. One symlink-specific test is skipped because symlinks are unavailable on this Windows environment. This confirms the verifier contract/report codec, transaction receipt, lifecycle, journal, phase-evidence and rollback focused tests. It does not establish native phase-verifier integration or operational Apply/Restore. PREPARED remains diagnostic-only; other nine verifier contracts remain unimplemented. Gate 2 remains closed.

## 2026-11-25 — PREPARED phase-report adapter tests pending

New adapter: Python/rehab_excel/daily_input_prepared_phase_adapter.py. New tests: Tests/test_daily_input_prepared_phase_adapter.py. They cover native readback invocation, exact snapshot hash binding, report codec round-trip, evidence-root containment, no-overwrite and fail-closed status pending trusted attestation. NOT YET RUN. The confirmed 110 passed, 1 skipped focused run predates this addition and must not be represented as covering these new tests.

## 2026-11-25 — Full CI regression including PREPARED report adapter

GitHub Actions full `Tests` regression for commit `a559a35c9a618b0d83764349d3aaa48e0e9cd211`:

- Windows / Python 3.11: 2734 passed, 71 warnings, 34.45s, job success.
- Ubuntu / Python 3.11: 2734 passed, 69 warnings, 30.31s, job success.
- Both synthetic Python safety jobs succeeded.

Workflow uses `pytest -q --disable-warnings --maxfail=35`; warning counts are visible, individual warning details are suppressed. This run includes the new PREPARED phase-report adapter and its tests. Latest code/test head `a8b1c59d5ec7faabfb8cbbce41bd80cc09582930` also completed full regression successfully: Windows 2734 passed / 71 warnings / 35.24s; Ubuntu 2734 passed / 69 warnings / 18.87s. Both synthetic safety jobs succeeded. Subsequent commits are documentation-only. User-local focused baseline before adapter: 110 passed, 1 symlink-specific skip. Latest adapter-specific Windows local test run is still required.


## 2026-11-25 — Offline internal verifier source audit

Added `Documentation/DAILY_INPUT_INTERNAL_VERIFIER_REVIEW_20261125.md` after reviewing the phase evidence policy, report/receipt codecs, existing INPUT_AND_LEDGER OOXML readback, PREPARED native COM pattern, applied archive, replacement facts and rollback audit.

New coverage gaps recorded:
- Transition evidence bundle completeness for phases requiring multiple EvidenceKind values.
- Receipt validation must derive required evidence internally, not accept a caller-selected subset.
- Trusted verifier/build-manifest attestation remains absent; hashes are integrity only.
- INPUT_AND_LEDGER requires an independent native Excel COM read-only proof reconciled with the existing OOXML proof and the same preflight facts.
- Explicit zero-change semantics are unresolved because the current OOXML readback requires non-empty validated changes.
- Phase-specific downstream verifier scopes and failure/crash-path test requirements are documented.

No tests or Excel were executed. User's latest local result (before commits `1f4a6c7` and `8179412`): 109 passed, 1 symlink skip, 17 warnings. Latest warning-related commits await local rerun. No zero-warning claim. Gate 2 remains closed.

## 2026-11-25 — Transition evidence bundle validator: full regression green

Added `daily_input_transition_evidence.py` and its focused tests. The bundle validator derives required evidence from the previous/candidate lifecycle transition; enforces exactly one report and receipt per required EvidenceKind; rejects missing, duplicate and extra kinds; binds all reports/receipts to one operation/revision/phase; cross-checks artifact/verifier metadata; checks the ordered receipt predecessor chain; and verifies exact artifact bytes.

Receipt binding was hardened: `validate_receipt_binding()` now requires the previous operation and derives the evidence obligations internally. The caller-controlled `required_evidence` parameter was removed.

GitHub Actions full regression for head `ed3bdec7e3e22a69b1162867f346321503b26498`:
- Windows: 2740 passed, 72 warnings, 44.59s.
- Ubuntu: 2740 passed, 68 warnings, 28.33s.
- Both synthetic safety jobs succeeded; overall workflow success.

The complete-bundle positive synthetic test uses temporary in-memory IMPLEMENTED contract statuses only to exercise the structural validator. Actual repository verifier statuses remain PREPARED=REPORT_ADAPTER_ONLY and nine NOT_IMPLEMENTED. This does not authenticate verifier origin, establish trusted attestation, authorize journal append or enable Apply. No native Excel COM/real XLSM test was run. Warning details remain suppressed by CI and counts are not zero. Gate 2 remains closed.

## 2026-11-25 — Trusted verifier attestation preflight: fail-closed, not authenticating

Added a strict untrusted attestation claim parser and non-authorizing reject-only boundary in `daily_input_verifier_attestation.py`, synthetic tests in `test_daily_input_verifier_attestation.py`, and a threat-model/key-custody design in `Documentation/DAILY_INPUT_TRUSTED_VERIFIER_ATTESTATION_DESIGN_20261125.md`. It checks claim envelope format, report/source/operation identity and malformed signatures; even a correctly shaped claim cannot pass the boundary because no external trusted key/build verification exists. No signing, real signature checking, trusted root, journal authorization or positive Excel proof was added. CI execution result must be recorded separately after final workflow review. Gate 2 remains closed.

### CI confirmation for attestation claim contract — 25/11/2026

GitHub Actions workflow `45598412259` for code head `a66088e91792b49095559bccdf9c800ede2ad7b0` completed SUCCESS:
- Windows / Python 3.11: **2756 passed**, 68 warnings, 44.39s.
- Ubuntu / Python 3.11: **2756 passed**, 68 warnings, 30.26s.
- Both synthetic safety jobs: success.

The workflow uses `--disable-warnings`; warning details are suppressed but warning totals remain visible. No native COM was run. These are tests of fail-closed metadata handling, not proof of cryptographic attestation or operational Apply/Restore. The trusted verifier signer/key custody/build pinning remain unimplemented; all claims are rejected by the current attestation boundary. Gate 2 stays closed.


## 2026-11-25 — INPUT_AND_LEDGER native readback synthetic baseline

Added `daily_input_input_ledger_native.py` and `test_daily_input_input_ledger_native.py`. Coverage includes: successful cross-view readback, exact saved input/ledger membership, native COM opened ReadOnly with macros disabled, disabled events/links/alerts, no Save, close/quit; failures for date, source reason, ledger identities, duplicate ledger entry, formula, visibility, protection, non-readonly Excel and cleanup failures, missing preflight, and OOXML rejection before COM launch. All examples use fictional synthetic data and fake COM objects.

Ubuntu Python 3.11 full CI at `6e4d99296fc6679fbaeffd4a2003a17ac165719f`: 2770 passed, 77 warnings, 24.10s. Windows outcome not yet confirmed at this documentation checkpoint. No native real Excel execution. Registry INPUT_AND_LEDGER remains NOT_IMPLEMENTED. Gate 2 closed.

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

## 2026-11-25 — PROJECTIONS_REOPEN partial structural baseline

New source and tests: `Python/rehab_excel/daily_input_master_projection_readback.py`, `Tests/test_daily_input_master_projection_readback.py`. Compares all canonical PATIENT_PLANNER formula links, optional displayed doctor identity, 11/13-column schemas, unique header, blank clinic separators/tail and no legacy N:P residue, with SHA-256 before/after. Reuses existing MasterProjectionRow contract, not a new scheduling interpretation. Synthetic corruption cases include wrong row/order/ID/formula/status, unapproved literals, missing headers, stale cells and missing sheet.

Not covered: recalculated cell values, therapist schedule projections, linkage to applied DAILY_INPUT/ledger, native Excel COM, signing/provenance, verified receipts or operational transaction. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED. CI pending verification at documentation checkpoint. No real CC workbook writes.


### MASTER structural readback CI defect and fix — 2026-11-25

The first synthetic regression at code/test commit `0142629049c37dce62c396e55abf0175e6576d88` failed 3 new MASTER positive tests on Ubuntu (2797 passed, 3 failed, 96 warnings). The diagnosis was incorrect literal comparison against Greek text after Unicode `casefold()`: the terminal sigma in `Θάλαμος` and `Ασθενής` was normalized to `σ`, while the comparison constants had `ς`. Fixed both comparisons by casefolding the canonical header constants too in code commit `905418df1c33b4008fd2039a89b6c06d7b4ea7fd`. New Ubuntu Python 3.11 CI `45617783680` passed **2800 tests / 99 warnings / 25.97s**, and both synthetic safety jobs succeeded; full Windows result to be confirmed separately before full green. No real Excel run or workbook modification. This was a newly introduced test/code mismatch, not evidence of operational workbook correction.


### Final cross-platform regression for MASTER structural readback — 2026-11-25

GitHub Actions run `45617783680` for corrected code commit `905418df1c33b4008fd2039a89b6c06d7b4ea7fd`: **SUCCESS** overall, all four jobs successful.
- Ubuntu Python 3.11: **2800 passed, 99 warnings, 25.97s**.
- Windows Python 3.11: **2800 passed, 93 warnings, 42.87s**.
- Both synthetic safety jobs passed.

This supersedes the provisional Windows-pending status for this run. Initial three failing new tests were corrected, not ignored. Warning details are suppressed by CI, counts are retained. This is saved-OOXML verification and synthetic tests only. Full PROJECTIONS_REOPEN remains NOT_IMPLEMENTED and no native Excel, therapist projection, authenticated attestation, Apply or Gate 2 action occurred.

## 2026-11-25 — THERAPIST_DAILY occupied distinct-clock OOXML diagnostic baseline

Files: `Python/rehab_excel/daily_input_therapist_projection_readback.py`; `Tests/test_daily_input_therapist_projection_readback.py`. Tests use fictional workbook cells and explicit preflight provider/clock expectations; check both blocks, distinct clock reuse, missing/extra provider, missing/extra slot, orphan occupancy, formula and invalid time, duplicate/incorrect expected facts, no write. This audit does not prove correct patient/day/session projections, formula calculations, therapist capacity or native Excel behavior. PROJECTIONS_REOPEN NOT_IMPLEMENTED. GitHub Actions full regression to be recorded after completion.


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

## 2026-11-25 — Read-only OOXML VBA-archive retention cleanup; local rerun pending

User's focused Windows run after the new native INPUT_AND_LEDGER and projection diagnostics: 214 passed, 1 symlink-specific skip, 40 PytestUnraisableExceptionWarning instances from `ZipFile.__del__` / `ValueError: I/O operation on closed file`. No test failures.

Source inspection found unnecessary `keep_vba=True` in read-only paths and synthetic non-VBA XLSM fixtures. Changed to `keep_vba=False` in:
- `daily_input_reader.py` (read-only reader, closes without saving)
- `daily_input_applied_state.py` (saved OOXML readback, closes without saving)
- `prepare_daily_input_journal_prepared.py` (PREPARED preflight inspection, closes without saving)
- `daily_input_master_projection_readback.py` (read-only structural verifier)
- `daily_input_therapist_projection_readback.py` (read-only structural verifier)
- the MASTER and THERAPIST_DAILY synthetic test mutation fixtures.

This does not write or save a workbook and does not alter VBA, workbook content, native COM security, transaction state or Apply/Restore behavior. The real Excel COM verifiers remain separate and unchanged. Source changes are committed; warning removal is NOT YET VERIFIED. Repeat the focused 13-file pytest command on Windows and report the actual warning count. Do not suppress warnings or infer a clean result from the code change alone.


## 2026-11-25 — Focused DAILY_INPUT evidence-chain Windows verification

Operator-confirmed local focused suite after synchronizing the current branch:

- **214 passed, 1 skipped in 12.34s**.
- No pytest warnings summary was emitted; **zero warnings were reported** for this exact 13-file run.
- The one skip is `Tests/test_daily_input_transaction_receipt.py:215`, because symlink creation is unavailable on the Windows platform.
- This confirms that the previously reported 40 `ZipFile.__del__` / `PytestUnraisableExceptionWarning` events are no longer reproduced in the focused suite after the targeted read-only `keep_vba=False` cleanup.

Scope: focused Python tests only. This is not a full-repository warning count, native Excel/COM transaction proof, signed verifier attestation, operational Apply/Restore authorization or Gate 2 acceptance. PREPARED and INPUT_AND_LEDGER remain REPORT_ADAPTER_ONLY; other phase verifier implementations and trusted attestation remain incomplete. Gate 2 remains closed.

 
## 2026-11-25 — Native MASTER projection diagnostic source regression

Added `daily_input_master_projection_native.py` plus synthetic fake-COM tests. Full GitHub Actions run **45657152050** for HEAD `e7ce7526de2c000cff9410819267f73fdc6ffc79` completed SUCCESS:
- Windows Python 3.11: **2848 passed, 35 warnings, 48.64s**.
- Ubuntu Python 3.11: **2848 passed, 34 warnings, 28.83s**.
- Both synthetic safety jobs passed.

The new diagnostic tests the saved OOXML structure and then uses an injected COM fake to test read-only native Excel lifecycle and evaluated MASTER room/name/status matching. It is not an actual Excel execution. This adds only a partial MASTER identity diagnostic; native therapist assignment extraction and complete projection/source-ledger proof remain pending. `PROJECTIONS_REOPEN` remains `NOT_IMPLEMENTED`; no report adapter or authorization. Warning counts remain nonzero.

 
### 2026-11-25 — Therapist AssignmentID verifier design source map

Source-only review mapped the stable identity chain through `stable_recurring_ids.py`, `reader.py`, `outpatient_schedule_source.py`, `materialize_sessions_for_date()`, `DAILY_REPLACEMENT_FACTS` and `build_daily_session_states()`. THERAPIST_DAILY does not store AssignmentID in its display cells; future observed membership must be reconstructed only by an unambiguous join from native-captured patient/provider/time cells to independently derived effective assignments. Ambiguity, missing joins, or unsupported student/zero-day semantics must fail closed. This is design/source evidence only, not a test or verifier implementation. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED.

 
### 2026-11-25 — Native THERAPIST_DAILY full-grid diagnostic CI

GitHub Actions run `45658828780`, code/test HEAD `71ff839699782615e9968d4213c9b8d42ea90bc6`, SUCCESS across all four jobs:
- Windows: **2859 passed, 32 warnings, 44.28s**.
- Ubuntu: **2859 passed, 35 warnings, 19.52s**.
- Both synthetic safety jobs passed.

The new diagnostic requires all 126 expected patient-grid cells and all 18 provider headers, then checks the native Excel read-only grid and time labels. Tests use fake COM only. It is not a complete assignment-ID or phase verifier; the expected-grid builder, AssignmentID membership extraction and operation/source-ledger binding remain pending. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED.

 
### 2026-11-25 — Expected THERAPIST_DAILY grid builder cross-platform regression

GitHub Actions run `45660810319`, HEAD `5a91058f0c9f78f63c19a70366987d8685796af7`, SUCCESS:
- Windows: **2866 passed, 36 warnings, 42.21s**.
- Ubuntu: **2866 passed, 32 warnings, 23.47s**.
- Both synthetic safety jobs passed.

New builder returns an explicit expected value for all 126 patient cells and expected active therapist AssignmentID facts using the existing patient-centric renderer. It includes synthetic absence, therapist replacement, student replacement, duplicate-display-name and identity-collision cases. Student rendering defect corrected. This is synthetic regression; no actual Excel COM execution. Full source capture/ledger linkage and operation-bound PROJECTIONS_REOPEN adapter remain pending.

 
### 2026-11-25 — Composite DAILY_INPUT projection evidence regression

GitHub Actions run `45662463260`, HEAD `6b55d2242b7ba1d3ec01ce01f5affa235717f94c`, SUCCESS:
- Windows Python 3.11: **2877 passed, 33 warnings, 45.11s**.
- Ubuntu Python 3.11: **2877 passed, 35 warnings, 18.84s**.
- Both synthetic safety jobs passed.

New composite binds prior INPUT_AND_LEDGER native readback, MASTER preflight source digest, full expected therapist grid digest and fresh native MASTER/therapist readbacks to the same workbook bytes/date. Synthetic tests reject cross-file/date/count/metadata mismatch and workbook mutation. Fake COM only; no real Excel. This is not a phase adapter or trusted attestation. PROJECTIONS_REOPEN remains NOT_IMPLEMENTED.

 
### 2026-11-25 — Blank THERAPIST_DAILY provider slots

GitHub Actions run `45663881144`, HEAD `46d4023b0ba8e5c65510fd1ddc5bb249fdadb3ce`: Windows **2880 passed / 34 warnings**; Ubuntu **2880 passed / 32 warnings**; both synthetic safety jobs passed. Added coverage for unused blank provider columns and rejection of occupancy beneath a blank header. All 18 header positions and 126 patient cells are still checked. Fake COM only.

 
### 2026-11-25 — DAILY_INPUT preflight/native wrapper regression

CI run `45667684716`, HEAD `36e7e1d3646d378a49c5acfe92b80fd2acac6ca2`, SUCCESS:
- Windows Python 3.11: **2897 passed, 29 warnings, 48.44s**.
- Ubuntu Python 3.11: **2897 passed, 32 warnings, 19.99s**.
- Both synthetic safety jobs passed.

Covers synthetic preflight assembly, shared-clock exact approval membership, native wrapper orchestration/hash/date/count drift, and safe aggregate-only CLI. Fake COM and synthetic data only. No real Excel/operational workbook. Not a phase adapter or trusted attestation.

 
### 2026-11-25 — Native source-sheet capture regression

CI run `45674827673`, HEAD `c9f7413c7ec409327a76a1a493a66dead0c22765`, SUCCESS:
- Windows Python 3.11: **2907 passed, 30 warnings, 46.67s**.
- Ubuntu Python 3.11: **2907 passed, 30 warnings, 17.42s**.
- Both synthetic safety jobs passed.

Covers typed OOXML cell parsing, native literal source capture, formula/error rejection, fixed UsedRange bounds, source hash immutability, mandatory sheet scope and full wrapper integration. Fake COM only. No real Excel.

 
### 2026-11-25 — Expanded native MASTER formula/evaluated-value regression

CI run `45676238896`, HEAD `1ebf531b276bf11960f21153e92db0fa16392bbc`, SUCCESS:
- Windows Python 3.11: **2911 passed, 30 warnings, 45.81s**.
- Ubuntu Python 3.11: **2911 passed, 33 warnings, 29.60s**.
- Both synthetic safety jobs passed.

Covers complete 10-column formula linkage plus status, 11 evaluated cells per row, infectious/room/identity/status value checks, doctor literal identity, formula error rejection and composite evidence count binding. Fake COM only.

 
### 2026-11-25 — PATIENT_PLANNER formula bridge / MASTER source-value equality

CI run `45677349897`, HEAD `6506354e2c4d38873ccf4707564ae2327dfc6854`, SUCCESS:
- Windows Python 3.11: **2916 passed, 30 warnings, 34.64s**.
- Ubuntu Python 3.11: **2916 passed, 32 warnings, 28.98s**.
- Both synthetic safety jobs passed.

Covers direct PATIENTS A:E formula semantics in PATIENT_PLANNER, native formula result equality, full MASTER formula outputs compared to PATIENT_PLANNER source cells, and independent infectious/room/identity/status checks. Fake COM only.

 
### 2026-11-25 — Guarded planner formula / expanded MASTER schema

CI run `45679125582`, HEAD `2b9c6294232e476216f8438e2d1880797a539b0d`, SUCCESS:
- Windows Python 3.11: **2922 passed, 30 warnings, 42.15s**.
- Ubuntu Python 3.11: **2922 passed, 31 warnings, 29.86s**.
- Both synthetic safety jobs passed.

Covers exact blank-guarded PATIENT_PLANNER A:E formulas, whitespace-vs-empty semantics, integrated OOXML/native capture, and expanded 13-column MASTER status-M readback. K/L remain outside proof. Fake COM only.

 
### 2026-11-25 — PATIENTS / PATIENT_PLANNER identity mirror regression

CI run `45680015176`, HEAD `fbf6e5571ad7cf6be412f719f304ade5d3c1322c`, SUCCESS:
- Windows Python 3.11: **2923 passed, 32 warnings, 44.62s**.
- Ubuntu Python 3.11: **2923 passed, 29 warnings, 24.37s**.
- Both synthetic safety jobs passed.

Covers A:E row-by-row identity mirror, formula/literal distinction, text/numeric preservation, wrapper patient-count binding and mismatch rejection. Fake COM only.


### 2026-11-26 — PREPARED_SNAPSHOT digest mismatch regression added

- Test: `Tests/test_daily_input_phase_verifier_contract.py::test_prepared_snapshot_report_rejects_digest_mismatch`.
- Scope: synthetic negative contract test; mismatched `prepared_snapshot_sha256` must not equal the exact saved evidence artifact `evidence_sha256`.
- Expected: fail closed before any positive-phase verifier availability/authorization decision.
- Code path: PREPARED contract equality guard; PREPARED report adapter binds both digests to the same saved XLSM bytes and checks byte immutability during readback.
- Execution: **PENDING for this one newly added test only**. The prior 13-file Windows focused suite is already GREEN at 214 passed / 1 symlink-specific skip / no warnings summary and includes both phase-report adapters; do not repeat it. When the operator returns to the PC, run only `Tests/test_daily_input_phase_verifier_contract.py::test_prepared_snapshot_report_rejects_digest_mismatch -q -rs`. No PC command during mobile-only mode.
- Native Excel / operational acceptance: NOT ESTABLISHED. This test uses synthetic report data only.
- Transaction authorization / Apply / Restore / Gate 2: unchanged; Gate 2 CLOSED.

### 2026-11-26 — Projection chain inventory correction and MASTER K/L coverage gap

- Existing implementation inventory: native source-sheet capture; PATIENTS/PATIENT_PLANNER A:E identity mirror; complete therapist expected-grid builder (126 cells); canonical date-specific source preflight; one-call native source + INPUT_AND_LEDGER + MASTER + THERAPIST_DAILY readback; exact workbook digest/date/count composition.
- Latest recorded full cross-platform synthetic CI: run \`45680015176\`, HEAD \`fbf6e5571ad7cf6be412f719f304ade5d3c1322c\`; Windows 2923 passed / 32 warnings; Ubuntu 2923 passed / 29 warnings; both synthetic safety jobs passed. Fake COM only, not real Excel.
- No repeat of those already-green tests absent a relevant code change.
- Remaining phase gap: \`PROJECTIONS_REOPEN\` still has no report adapter/registered verifier, trusted operation-bound attestation or authenticated shared-slot approval source; remains NOT_IMPLEMENTED.
- Newly confirmed mandatory application gap: MASTER K (psychology) is rendered by the existing patient-program VBA refresh path but is cleared/not repopulated by bulk \`master_projection_writer.py\`; MASTER L (afternoon programme) has a separate \`AFTERNOON_PHYSIO_SLOTS\` source contract and read-only patient panel but no MASTER renderer or operational persistence/history acceptance. The current expanded MASTER test explicitly expects K/L blank, so it is not coverage of the required business content.
- K/L are not authorized deferrals. New focused synthetic tests must be added only alongside the actual K/L integration, covering source-to-PatientID-to-MASTER mapping, blank PM default, exact day/time/provider rendering, stale-row clearing and no AM/PM cross-contamination. Native Excel proof remains PC-only.

### 2026-11-26 — K/L display contract audit

- K psychology display is defined by the existing VBA renderer: \`Ψυχ_Ώρα\`, \`Ψυχ_Ημέρες\`, \`Ψυχ_Ψυχολόγος\`, in that order, non-empty parts newline-separated; recognized date/numeric time becomes \`hh:mm\`. Bulk writer parity can be implemented and tested synthetically.
- L source schema is explicit and independent, but no MASTER display convention is approved. Patient-panel \`PM_ROW\` is a transport format, not cell formatting policy. Do not guess grouping, ordering or duration. Keep L empty until exact convention is confirmed.
- No code changed in this audit. Native Excel remains PC-only.

### 2026-11-26 — Afternoon PM implementation deferred pending PC visual review

- Do not implement MASTER L display/layout until reviewed on PC in a disposable DEMO workbook; the earlier line-format suggestion is not approved.
- PM patients are a subset of AM patients and must retain the same PatientID. Afternoon days are selected explicitly; reuse the existing morning weekday/day-combination rules. Never assume daily attendance or create a separate PM patient registry.
- General treatment duration is 45 minutes. Existing PM source stores StartTime only; no end time/duration is persisted. Do not fabricate an end time or silently change the source schema.
- No code/workbook/test changed for this user clarification. No repeat tests.

### 2026-11-26 — PM requirements clarified; implementation deferred to PC review

- No afternoon MASTER L renderer/layout work until the user can inspect the actual disposable DEMO at a PC. No guessed cell display convention.
- Afternoon patients are a subset of morning patients; preserve the same PatientID identity.
- PM sessions are optional on selected weekdays. Reuse the existing morning weekday selection and valid day-combination rules; do not default to all weekdays.
- General session duration is 45 minutes, but the current PM sheet stores start time only. No end-time inference or schema expansion before the PC design review.
- No PM runtime tests are warranted until code changes are made. Existing PM source/schema tests are already recorded; do not repeat them.

### 2026-11-26 — MASTER K writer change; full CI pending

- Changed: expanded MASTER writer now renders psychology K from PATIENT_PLANNER exact headers `Ψυχ_Ώρα`, `Ψυχ_Ημέρες`, `Ψυχ_Ψυχολόγος`, matching existing VBA component order and omission behavior. Recognized time values render as `hh:mm`; unknown text remains literal.
- Fail-closed identity check: PATIENT_PLANNER row PatientID must match the projection PatientID before K is written.
- Saved-output verification checks K against source values. L remains deliberately blank; PM layout/scheduling deferred to PC review.
- Tests changed: `Tests/test_master_projection_writer.py`, including expanded-schema K output and a focused display helper regression. No local execution claimed.
- GitHub Actions run `45870586916`: synthetic safety jobs passed on Windows and Ubuntu; full Windows/Ubuntu Python regression jobs were still running at checkpoint. Record exact totals only after completion.
- No native Excel proof, no VBA change, no operational workbook change. Do not repeat previous focused suite; this change triggered the full current regression workflow.

### 2026-11-26 — MASTER K implementation final cross-platform CI result

- Run: `45870877451`; code HEAD: `7799df57f9dc916fec7acb8eaa66ac4725297531`.
- Complete Python regression: **Windows 2926 passed / 35 warnings; Ubuntu 2926 passed / 29 warnings**.
- Synthetic safety jobs: GREEN on both platforms.
- Includes expanded MASTER psychology K rendering, saved-output readback, empty-component/time formatting checks and K/L stale-row-boundary test.
- An initial run failed two new tests due to a missing `datetime` import; corrected, then full cross-platform regression passed. Preserve this history; do not count the initial attempt as GREEN.
- No native Excel proof, no VBA or production workbook change. K still needs disposable native Excel verification on PC. PM L/UI and day-combination work remain deferred until PC review.

### 2026-11-26 — MASTER K independent preflight/readback binding final regression

- Run: `45872223262`; tested HEAD: `e1c8cab22f7039c62ed31d57a80978851534bf91`.
- Full regression: **Windows 2934 passed / 30 warnings; Ubuntu 2934 passed / 33 warnings**.
- Synthetic safety jobs: GREEN on both platforms.
- New end-to-end expected-value binding: `build_master_projection` derives psychology K by PatientID from PATIENT_PLANNER; writer verifies live source identity/value before write; saved OOXML and native readback compare K to the same independent expectation and require L blank.
- Expanded test fixtures now include a psychology expectation on every projected row. Earlier incomplete-fixture failures were fixed; retain the failure history.
- No native Excel COM or operational workbook proof; no VBA change. K still requires disposable native Excel verification. PM L/UI remains deferred pending PC review.

### 2026-11-26 — Psychology source header and PatientID binding regression

- Run: `45872660668`; tested HEAD: `1f0b57ff7c578475f50383d7bcdb212caa876aea`.
- Full regression: **Windows 2935 passed / 28 warnings; Ubuntu 2935 passed / 30 warnings**.
- Synthetic safety jobs: GREEN on both platforms.
- Expanded MASTER requires unique PATIENT_PLANNER headers for PatientID and all three psychology source fields. Projected inpatient PatientID must match the planner row before K is derived. Outpatients are intentionally excluded from this new row-presence requirement.
- New synthetic coverage: missing psychology header, duplicate psychology header, PatientID row drift.
- No native Excel COM proof, no VBA or production workbook change. Native disposable XLSM check still pending.

### 2026-11-26 — Afternoon edit review uses existing 45-minute overlap semantics

- Run `45872972784`; tested HEAD `a81a70cc7d225995267af369376a6b2d0358c237`.
- Full regression: **Windows 2936 passed / 28 warnings; Ubuntu 2936 passed / 30 warnings**.
- Synthetic safety jobs: GREEN on both platforms.
- `review_afternoon_edit_plan` now checks the complete proposed schedule with the existing scheduler's 45-minute ΦΘ overlap function, in addition to exact shared-slot approval. Staggered same-therapist overlap and same-patient overlap fail; adjacent slots pass.
- Pure/read-only policy only. No persistence, UI, VBA or native Excel operational proof.

### 2026-11-26 — Afternoon weekday-combination parity regression

- Run `45874054596`; tested HEAD `04e2d71b79281625b5d16bb20efcc9a5d7646e20`.
- Full regression: **Windows 2938 passed / 30 warnings; Ubuntu 2938 passed / 29 warnings**.
- Synthetic safety jobs GREEN on both platforms.
- New coverage: PM draft parser accepts multiple explicit weekday groups `ΔΕ-ΤΡ/ΠΕ-ΠΑ` with one explicit time/provider and rejects a weekday repeated across groups.
- Uses the shared `parse_weekday_groups` from morning mapping. No UI, PM persistence, MASTER L display, workbook or VBA changes. No native Excel/operational proof.

### 2026-11-26 — PROJECTIONS_REOPEN report adapter full regression

- Run `45875449647`; tested HEAD `1bc189a4903fea1890f82efee9555546bf587eac`.
- Full regression: **Windows 2946 passed / 30 warnings; Ubuntu 2946 passed / 30 warnings**.
- Synthetic safety jobs: GREEN on both platforms.
- Added report adapter for legal INPUT_AND_LEDGER_COMMITTED -> PROJECTIONS_VERIFIED transition. Reuses composite native diagnostic for source capture, input/ledger, MASTER and complete therapist grid; binds the report to exact post-commit evidence bytes, service date and preflight scope.
- Synthetic negative coverage: source/evidence hash distinction, wrong digest/version/scope, workbook mutation, path confinement, output collision and phase identity.
- Status REPORT_ADAPTER_ONLY; no trusted attestation, no receipt/journal authorization, no real native Excel or production workbook execution.


### 2026-11-28 — Operator PowerShell execution incident (not a test result)

- The three-candidate H10 K schema comparison did not complete. The pasted wrapper split across interactive submissions; finally was treated as a standalone command and the active transcript blocked its own rename.
- Classification: **wrapper SAFETY STOP; schema comparison PENDING**. The printed FINAL_STATUS=B is not valid evidence and must not be used to mark any workbook candidate BAD.
- No workbook write/native Excel action occurred.
- Mandatory future method: one complete .ps1 executed once, or one complete & { ... } pasted in a single submission. Never split compound clauses or here-string delimiters. Stop transcript before moving its file; check $LASTEXITCODE immediately. See Documentation/POWERSHELL_OPERATOR_EXECUTION_STANDARD.md and guide §4.253.
- This incident adds no pass/fail count and does not invalidate earlier green focused results.


### 2026-11-28 — Pending replacement inventory snapshot builder

- Code/test HEAD: \`bf8c6e34c71c45208f8312ff857aafe698cff07c\`.
- GitHub Actions run: [46161976361](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/46161976361), SUCCESS.
- Full regression: **Windows 2958 passed / 28 warnings; Ubuntu 2958 passed / 27 warnings**.
- Synthetic safety jobs: SUCCESS on Windows and Ubuntu.
- Added canonical read-only inventory derived from exact preflight states; includes every/only therapist-absent pending target, validates exact base-session identity joins, counts all five states, binds source and evidence hashes separately, and writes an exclusive root-confined artifact.
- This is not native Excel proof, not a phase-report adapter, not a workbook write and not transaction authorization. \`PENDING_DECISION_INVENTORY\` remains NOT_IMPLEMENTED.


### 2026-11-28 — PENDING_DECISION_INVENTORY phase report adapter

- Code/test HEAD: \`1b90e3dab9e5dc3d4bd9166e5c321ab665495cde\`.
- GitHub Actions run: [46162899470](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/46162899470), SUCCESS.
- Full regression: **Windows 2964 passed / 30 warnings; Ubuntu 2964 passed / 28 warnings**.
- Synthetic safety jobs: SUCCESS on Windows and Ubuntu.
- Added strict canonical inventory artifact decode/readback and a phase-report adapter. The adapter independently rebuilds the expected inventory, checks artifact bytes/digest, operation/date/source/workbook binding, reuses the existing composite native read-only verifier, checks native version/readback scope and publishes the report exclusively.
- Status: **REPORT_ADAPTER_ONLY**. CI uses fake native result injection. Actual Excel COM/native acceptance, trusted attestation, receipt/journal append and Apply authorization are NOT established.
- One earlier run had 1 stale registry-test expectation failure; the test was corrected and the final exact head is GREEN.
- Next: REPLACEMENTS_REOPEN report adapter; no replacement decision logic.


### 2026-11-28 — REPLACEMENTS_REOPEN report adapter

- Code/test HEAD: \`ad18f3305bd3d5f3a3caa7c1f6b55df575940f5a\`.
- GitHub Actions run: [46165623512](https://github.com/diomil73/EKA-REHAB-TIMETABLE/actions/runs/46165623512), SUCCESS.
- Full regression: **Windows 2972 passed / 29 warnings; Ubuntu 2972 passed / 28 warnings**.
- Synthetic safety jobs: SUCCESS on Windows and Ubuntu.
- Added canonical evidence builder/reader and a phase-report adapter for replacement facts and treatment suspensions. Every pending inventory target must resolve exactly once as replacement or explicit suspension; exact session/patient/original-provider/date joins and effective daily state are checked.
- Reuses existing macro-disabled native source-cell capture for DAILY_INPUT and DAILY_REPLACEMENT_FACTS; exact workbook digest and native Excel version are bound.
- Status: **REPORT_ADAPTER_ONLY**. CI native capture is injected/fake. Actual Excel COM, trusted attestation, OPERATOR_RESOLUTION authentication, receipt/journal append and Apply authorization are NOT established.
- Next: separate OPERATOR_RESOLUTION evidence artifact/reader; then integrate both required reports.

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
