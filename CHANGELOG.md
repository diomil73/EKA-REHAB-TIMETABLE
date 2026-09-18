## 2026-09-19 — Replacement visual semantics

- Replaced sessions no longer show the patient with strikethrough in the original slot.
- The original line is rendered in muted italics because the session still takes place.
- Strikethrough remains reserved for absence/cancellation-style states.
- Rich-text writeback now supports per-run italic formatting.

# Changelog

## Unreleased

### Added
- Replacement ranking based first on operational daily workload, then exact-time match, then another common free timeslot.
- Alternative-timeslot visibility for replacement candidates.
- Student replacement providers with active-placement and capability checks.
- Student daily capacity of 5 patient/timeslots.
- Student workload and exact-timeslot availability engine.
- Provider-kind metadata so therapist and student replacements remain distinct.
- Student double-booking conflict classification.

### Changed
- A provider busy at the requested time is no longer automatically removed from replacement suggestions when another common free timeslot exists.
- Infectious workload is now a tie-breaker after total load and exact-time availability, rather than overriding the confirmed load-first priority.

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
- Added safety gate that blocks operational import when authoritative identity/source errors exist.
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
