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
