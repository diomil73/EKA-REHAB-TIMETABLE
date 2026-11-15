## 2026-09-19 — Replacement visual semantics

- Replaced sessions no longer show the patient with strikethrough in the original slot.
- The original line is rendered in muted italics because the session still takes place.
- Strikethrough remains reserved for absence/cancellation-style states.
- Rich-text writeback now supports per-run italic formatting.

# Changelog

## Unreleased

### Added
- Application-style outpatient schedule manager listing patient, treatment, time, days and therapist while keeping worksheet row/internal target mapping hidden from the user.
- Outpatient daily-edit hardening that removes stale plain/decorated patient render lines from affected old/destination `THERAPIST_DAILY` cells before rendering the edited placement once.
- Regression coverage proving one outpatient recurring entry materializes to one operational session and contributes one occupied therapist timeslot.
- Replacement ranking based first on operational daily workload, then exact-time match, then another common free timeslot.
- Alternative-timeslot visibility for replacement candidates.
- Student replacement providers with active-placement and capability checks.
- Student daily capacity of 5 patient/timeslots.
- Student workload and exact-timeslot availability engine.
- Provider-kind metadata so therapist and student replacements remain distinct.
- Student double-booking conflict classification.
- In-session daily-ledger restore bridge that keeps the authoritative Python preview/recompute gate while applying the verified result back into the already-open workbook with a rollback copy.
- Backend restore profiler with Excel-output and pre-Excel phase timings for targeted latency work.

### Changed
- Outpatient smoke installation now upgrades legacy registration menus that did not yet contain an outpatient action: it creates the visible `Προγράμματα εξωτερικών ασθενών` button, moves the close action out of the way, and routes the new button to the schedule manager.
- `EditOutpatientSchedule` now opens the business-facing outpatient schedule manager instead of asking the user for an `OUTPATIENT_SCHEDULE` worksheet row. The existing registration menu outpatient action is routed to the same manager.
- Outpatient recurring edit carries `target_base_entry_id` through conflict validation and the final in-session write. The writer validates the target row and PatientID, updates the same `outpatient:<row>` instead of appending a duplicate, and refreshes the current-date therapist-daily placement when therapist/time/day changes.
- Native outpatient edit debugging identified a daily-projection representation mismatch: normal daily rendering may decorate the patient name (for example with `[Καθ/να]`) while the old in-session edit cleanup matched only the plain name. The fresh smoke builder now installs a post-conflict hardening patch so stale decorated lines are removed before the edited placement is rendered once.
- Locked the requirements for global multi-therapist-absence replacement planning: scarcity-aware assignment across all displaced patients, persistent pool/robotic operating hours, reclining-treatment adjacency preferences, workload-balance goals, and user-approved last-resort double-booking.
- Clarified that patient-specific physical suitability, compatibility, and waiting tolerance should be captured selectively rather than creating mandatory high-volume data entry.
- Audited the existing replacement path: candidate ranking remains workload-first and the workflow remains sequential; global scarcity-aware planning and persistent pool/robotic operating-hour constraints are documented as required follow-up work, not claimed as implemented.
- A provider busy at the requested time is no longer automatically removed from replacement suggestions when another common free timeslot exists.
- Infectious workload is now a tie-breaker after total load and exact-time availability, rather than overriding the confirmed load-first priority.
- Daily ledger restore no longer uses a visible `WScript.Shell.Exec` console or the detached workbook close/reopen commit worker on the primary UI path. Backend execution is hidden and blocking, then the verified operational sheet set is synchronized in-session and saved without closing the workbook.
- Native Excel correctness is now confirmed for per-row restore, selected multi-row restore, and restore-all, with exact fresh reconstruction/statistics checks after each scenario.
- User-facing ledger `Αιτία` now shows agreed Greek business wording for known technical replacement reasons instead of exposing internal English labels.
- Restore performance work reduced backend latency from about 45.7s to 31.3s by batching MASTER projection COM writes; profiling then isolated the remaining dominant pre-Excel phases.
- Restore recomputation now reuses the already-loaded patient registry and already-built MASTER projection instead of reparsing those identical inputs a second time.
- Selected/all restore no longer starts a second hidden Excel process only to mark `DAILY_REPLACEMENT_FACTS` as RESTORED. The requested set is validated in memory, effective ACTIVE replacements are derived without mutating the stage, and authoritative fact statuses are persisted in the existing verified-output Excel session before its single save.
- MASTER projection metadata now reads PATIENTS source rows, SETTINGS room order, and MASTER header position directly from the read-only OOXML ZIP/XML package instead of loading the full macro-enabled workbook through openpyxl.
- Restore latency remains an open performance/UX item until native total time reaches the acceptance target.

### Unchanged
- Base schedule remains immutable.
- `Rehab_Center_System_v27_1.xlsm` baseline remains untouched.

## 2026-09-17 - Read-only Excel integration

- Added read-only `.xlsm` adapter for `PATIENTS`, `PATIENT_PLANNER` and `SETTINGS`.
- Added recurring `BaseScheduleEntry` and date materialization into operational `Session` objects.
- Added support for the workbook day pattern `Καθ/να`.
- Added workbook audit tooling and regression tests.
- Confirmed `SESSIONS` identity drift and excluded it from the Python scheduling import path for now.

## 2026-09-18 - Authoritative source contract

- Formalized `PATIENTS`, `PATIENT_PLANNER` and `SETTINGS` as the only current authoritative Excel import sources.
- Added immutable read-only `WorkbookSnapshot`.
- Added a safety gate that blocks operational import when authoritative identity/source errors exist.
- Classified `SESSIONS` as audit-only and workbook output/UI sheets as non-authoritative.
- Added a privacy rule: no automatic patient-data export to repository files.
## 2026-09-18 - Excel write-back dry run

- Added validated `CellPatch` / `WritePlan` objects for future Excel output.
- Added a dry-run safety gate that permits only sheets marked as future write targets.
- Protected authoritative input sheets and legacy operational sheets from Python write-back.
- Added workbook SHA-256 verification proving that dry runs do not alter the `.xlsm`.
- Added CLI tooling and regression tests for the write-back safety layer.
- Real workbook writes remain disabled.


## 2026-09-18 - Excel layout contract

- Mapped the actual v27.1 cell structure for `MASTER_SCHEDULE`, `THERAPIST_DAILY`, `REPLACEMENTS`, `CONFLICT_LOG`, `REPLACEMENT_LOG` and `THERAPIST_ATTENDANCE`.
- Added a read-only OOXML package probe that inspects `.xlsm` layout without saving the workbook.
- Added dynamic `THERAPIST_DAILY` provider/timeslot cell resolution.
- Added cell-level write zones in addition to the existing sheet-level source contract.
- Protected therapist headers, time cells, refresh-button area and master-schedule identity/status columns from future write plans.
- Added a layout drift audit and regression tests.
- Real Excel mutation remains disabled.

## 2026-09-18 - Daily render planning
- Added `DailySessionState -> Excel cell` binding for `THERAPIST_DAILY` and `MASTER_SCHEDULE` references.
- Added semantic formatting roles for active, absent, replacement, infectious and robotic sessions.
- Preserved multi-patient cells and student legacy-column resolution.
- No workbook writes are performed by this layer.

## 2026-09-18 - Native Excel copy-write foundation

- Added rich-text `CellPatch` runs for line-level strike-through and font roles.
- Added `DailyExcelRenderPlan -> WritePlan` translation for `THERAPIST_DAILY`.
- Added Windows-native Excel COM writer that writes only to a new `.xlsm` copy.
- Added source SHA-256 verification and deletion of partial preview files on failure.
- Added semantic presentation handling for infectious, robotic, student and muted lines.
- Added a Windows/Excel runtime probe.
- Production/baseline in-place write-back remains prohibited.

- Added compact two-line replacement overlay in the original `THERAPIST_DAILY` slot.
- Original patient remains struck through; replacement provider and effective time appear directly below.
- Active student replacement names use the student-green presentation role.
## 2026-09-19 - Generic daily operational overlay

- Added a minimal overlay writer that touches only changed `THERAPIST_DAILY` cells and preserves unrelated existing lines.
- Added repeatable CLI inputs for replacements, patient absences and therapist absences.
- Replacements now keep the original patient in muted italics without strikethrough and append `→ provider HH:MM`.
- Absence/cancellation-style states retain strikethrough semantics.
- Added multi-replacement same-cell regression coverage and corrected `DailySessionState.original_should_be_struck_through`.


## 2026-11-12 — Native synthetic identity proof, read-only event history and complete handoff

- Verified by operator on actual Windows Excel: isolated read-only CC probe; fresh synthetic Save/close/read-only reopen with preserved VBA and unchanged source; one-shot native stable AssignmentID migration with **179 IDs** independently read back, source CC and synthetic history unchanged. All three A candidates frozen.
- Added effective-dated recurring change proposal and capacity/patient-collision checks, read-only Excel bridge and strict immutable `AssignmentID` binding.
- Added pure migration plan, one-shot **synthetic-only** native Excel identity installer, synthetic COM tests, and independent history/VBA readback.
- Added optional read-only `RECURRING_CHANGE_EVENTS` parser: validated immutable event IDs, assignment/patient match, effective-date chronology and duplicate prevention. No native event append, real treatment edit, rollback or production use yet.
- GitHub Actions 2026-11-12 run `43313086923`: **1,646 Python tests passed on each Windows and Linux**, and **182 targeted tests on Windows**. Those runs do not exercise actual Excel COM.
- Created `Documentation/AUTHORITATIVE_RESUME_20261112.md` as a complete current operating handoff and refreshed ONBOARDING, CURRENT_HANDOFF, PROJECT_STATUS, IMPLEMENTATION_MAP, DECISIONS and native evidence runbooks.
- Outstanding: synthetic event append/save/reopen with compensation and full capacity validation, real-data staging identity/conflict clearance, and management reporting visual/privacy acceptance. Production MASTER and real-patient files remain unmodified by these works.



## 2026-11-12 — Source-only permanent treatment event append proposal

- Added `Python/rehab_excel/recurring_change_append_plan.py`: non-mutating SHA256-bound append preflight, existing event uniqueness/strict chronology and immutable AssignmentID ownership, using current event history plus approved per-provider limits.
- Added isolated regression tests and Windows/Linux synthetic CI wiring. This does **not** write the event to Excel and does not implement native rollback or clinical import. Existing 179-ID synthetic native proof stays frozen.


### 2026-11-13 real-source demo preflight preparation (documentation-only)
See `Documentation/REAL_SOURCE_DEMO_WEEKEND_GATE_20261113.md`. Audited tracked read-only source gate and tests; created operator-safe PowerShell and stop/runbook. No real patient import, code changes, current staging check, native Excel proof or demo delivery was performed. Do not reuse protected CC or frozen candidates.


### 2026-11-13 source-only trial import planning checkpoint
Added `Python/rehab_excel/trial_import_plan.py` (commit `b0e9d48`) and `Tests/test_trial_import_plan.py` (commit `0a9bc4c`). The plan reuses strict read-only combined preflight and exact approved `SourceWeeklyEntry` admission, cross-checks counts, binds both source SHA256 values, rejects changed sources, and hides private entries from its object repr. **It does not write an XLSM, implement a native importer, or authorize clinical use.** The new test file has six synthetic cases, but **tests/CI have NOT yet been observed passing**. Before native integration run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` in one PowerShell fail-fast block; stop on red. No private patient data uploaded. Current real staging remains unverified. More work is needed to build the new isolated native importer and independent readback.


### 2026-11-13 verified-schema weekly projection SOURCE checkpoint
Source-only modules added: `Python/rehab_excel/trial_weekly_projection.py` (`f90d4b6`) and seven synthetic tests `Tests/test_trial_weekly_projection.py` (`f2e6268`). This pure bridge maps independently approved `TrialImportPlan` entries into existing `PROGRAM_WEEKLY_SLOTS` six-column headers, with strict inpatient ID membership, no silent omissions, treatment/provider checks and patient-time collision rejection. The synthetic roundtrip test invokes the **existing `read_weekly_slots_sheet`** against an in-memory workbook. No real patient data, Excel file output, COM write, native acceptance or approved clinical import yet. New tests are **not confirmed green** unless a subsequent CI/native result explicitly proves that. Run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_weekly_projection.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` as one fail-fast PowerShell gate. Still missing: approved patient registry detail mapping, a real source-bound native writer, independent readback, privacy and operator proofs. Do not treat `PROGRAM_WEEKLY_SLOTS` alone as a complete operational demo.


## 2026-11-15 — Evidence-based baseline matrix and remaining scope clarification

Created `Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md` after reviewing the H2-H8 acceptance checkpoints, H8 dashboard feedback, H10 therapist chart requirements, `STATISTICS_HISTORY_DESIGN.md`, `AUTHORITATIVE_RESUME_20261112.md`, and the current afternoon branch handoff.

Key distinction: most core daily mechanisms have substantial operator-confirmed native evidence (daily absence Apply/Restore, day close/amendment, therapist and patient session history, contextual day save, integrated reports). H8 dashboard artifact BL is manually accepted for its specifically recorded UI scope. This does **not** mean every analytics feature is complete. The patient 15-day progress/evaluation workflow remains design-only. H10 therapist charts/management have partial native/source evidence and later visual follow-ups; genuine therapist absence/working-day history is unavailable. Other explicitly open mechanisms remain global multi-absence optimization, permanent recurring-change native write/rollback/UI, real-source native importer/admission, and final worksheet protection classification. Afternoon read-only UI is source-green at 49/9.22s, but has no native Excel acceptance or save path.

The matrix records exact accepted/frozen artifact names and scope boundaries, and is the guide for selecting component references when building the single full-application synthetic test workbook. Frozen proof files are evidence only, never writable build targets. No workbook was opened or modified in this documentation review.


## 2026-11-15 — Full DIOMIL workbook candidate architecture audit

- Added `Python/tools/audit_integrated_workbook_structure.py` for non-identifying, read-only XLSM/XLSX structural inventory (sheet visibility/dimensions, formula/cache counts, defined-name count, ZIP/VBA package metadata and schema role hints).
- Added focused tests in `Tests/test_audit_integrated_workbook_structure.py` covering read-only hash preservation, no patient/staff cell values in output, hidden/veryHidden surfaces, invalid/corrupt inputs, scan-limit fail-closed behavior and multi-candidate execution.
- Updated `Documentation/CURRENT_HANDOFF.md`, `Documentation/PROJECT_STATUS.md`, and `Documentation/INTEGRATED_TEST_BASELINE_MATRIX_20261115.md` with the exact execution gates and candidate restrictions.
- The four H10 workbook files remain unselected and read-only. No workbook was opened in Excel or modified. The new focused tests have not yet been run by the operator.


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
