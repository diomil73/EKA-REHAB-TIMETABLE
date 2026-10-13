# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V71.xlsm`
- V71 now builds and opens after the earlier installer hardening.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use.
- V71 conflict behavior now appears functionally correct in native use: overlap is detected and Save presents the intended conflict decision flow.

## Conflict-validation issue found in V70

Native smoke exposed a consistency gap:

- a manual Λογο time was placed on top of ΦΘ on a shared day,
- **Αυτόματη συμπλήρωση** detected that something was wrong, but the message was too vague,
- **Αποθήκευση προγράμματος** accepted and persisted the conflict,
- MASTER then displayed the conflicting program normally.

The common conflict gate now addresses Auto and Save with detailed diagnostics and explicit override.

## Conflict-validation behavior

A common `validate-vba` bridge command checks pairwise patient-program overlaps using the duration contract:

- standard treatments: 45 minutes,
- Ανακλινόμενο: 30 minutes,
- only shared weekdays count,
- adjacent intervals are allowed.

Detailed diagnostics have the form:

`Δευτέρα: ΦΘ 09:00-09:45 συγκρούεται με Λογο 09:00-09:45.`

The patient-program UI:

- validates before **Αυτόματη συμπλήρωση**,
- validates before any atomic Save write begins,
- shows **Συνέχεια** and **Επιστροφή** on Save conflicts,
- **Επιστροφή** cancels before writes,
- **Συνέχεια** is an explicit override.

## V71 installer hardening already verified

Two installer blockers were found and fixed before native V71 could open:

1. VBE COM rejection of conflict-form `Width/Height`.
2. brittle exact marker for injecting `ValidatePatientProgram()` into `PatientProgramRuntime`.

The runtime patcher is now procedure-scoped, CRLF-safe and idempotent, while designer geometry writes are guarded.

Latest user-reported focused gate after those fixes:

`62 passed in 3.34s`

## New native V71 follow-ups

### Conflict dialog is too short

The Save conflict dialog opens, but its installed/default form geometry is too short. The full conflict text and the **Συνέχεια / Επιστροφή** buttons are not fully visible.

New fix implemented:

- `patient_program_conflict_layout_vba.py` replaces the conflict-form code after installation,
- the form now applies its own runtime layout in `Configure()`,
- runtime layout explicitly sizes the form, message label and both buttons,
- this avoids relying on VBE designer geometry persistence.

This runtime-layout fix is implemented but not yet user-test-verified.

### Existing Ρομποτικό time can still appear as decimal in MASTER

Native V71 showed the Ρομποτικό time for patient Γιαννιά in raw decimal form.

The likely cause is stale MASTER projection content for existing patients. `RefreshMasterPatientProgram()` already formats numeric times through `Format$(..., "hh:mm")`, but existing MASTER rows are not necessarily passed through that composite renderer during preview build.

New fix implemented:

- after the patient-program renderer is installed and verified, the preview installer iterates all nonblank PatientIDs in `PATIENTS`,
- it invokes `RefreshMasterPatientProgram` for each existing patient before the workbook is saved,
- this re-renders all current treatment cells as composite text and removes stale raw Excel time serials, not just the one observed Ρομποτικό cell.

This initial-all-patients MASTER refresh is implemented but not yet user-test-verified.

## Immediate next start point

1. Pull latest `feature/outpatient-scheduling`.
2. Run focused tests including:
   - `Tests/test_patient_program_conflict_layout_vba.py`
   - `Tests/test_patient_registration_scheduling_slice.py`
   - existing patient-program conflict/renderer tests.
3. If green, rebuild a fresh preview from clean V48 using the next preview version after V71.
4. Native re-test:
   - Save conflict dialog shows the full message and both **Συνέχεια / Επιστροφή** buttons,
   - patient Γιαννιά Ρομποτικό time appears as `hh:mm`, not decimal,
   - spot-check one or two other existing MASTER treatment cells for raw decimals.
5. Reconfirm conflict Return/Continue behavior and adjacency boundary.
6. Then continue with edit-existing-program and optional manual names.
7. Full regression suite only after native edit smoke is clean.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual day/time/provider constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration currently used by the solver: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Recliner adjacency to ΦΘ is a strong soft preference, not a hard rule.
- MASTER is a derived view; `PATIENT_PLANNER` remains authoritative for the inpatient patient program.
- Conflict override must be explicit. Returning for correction must perform no writes.

## Recovery order if context is lost

Read these in order:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. tests around the currently active patient-program modules

Use `python -m pytest` on Windows, never bare `pytest`.
