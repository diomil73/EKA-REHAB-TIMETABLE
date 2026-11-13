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
