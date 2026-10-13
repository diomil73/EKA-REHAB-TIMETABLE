# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V71.xlsm`
- V71 builds and opens after the earlier installer hardening.
- Next preview to build: `MASTER_TOOLBAR_SMOKE_V72.xlsm` from clean V48.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use.
- V71 conflict behavior appears functionally correct in native use: overlap is detected and Save presents the intended conflict decision flow.

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

User-reported focused gate after those fixes:

`62 passed in 3.34s`

## New V71 native follow-ups

### Conflict dialog too short

The Save conflict dialog opened, but its installed/default form geometry was too short. The full conflict text and the **Συνέχεια / Επιστροφή** buttons were not fully visible.

Fix implemented:

- `patient_program_conflict_layout_vba.py` replaces the conflict-form code after installation,
- the form applies its own runtime layout in `Configure()`,
- runtime layout explicitly sizes the form, message label and both buttons,
- this avoids relying on VBE designer geometry persistence.

### Existing Ρομποτικό time shown as decimal in MASTER

Native V71 showed the Ρομποτικό time for patient Γιαννιά in raw decimal form.

Fix implemented:

- after the patient-program renderer is installed and verified, the preview installer iterates all nonblank PatientIDs in `PATIENTS`,
- it invokes `RefreshMasterPatientProgram` for each existing patient before the workbook is saved,
- this re-renders all current treatment cells as composite text and removes stale raw Excel time serials, not just the one observed Ρομποτικό cell.

## Latest confirmed test checkpoint

Latest user-reported focused gate for the runtime conflict layout + existing-patient MASTER refresh batch:

`57 passed in 3.96s`

This confirms the source/test gate is green for:

- runtime sizing of the conflict dialog,
- initial refresh of existing MASTER patient-program rows,
- composite time rendering,
- scheduling-slice integration,
- existing patient-program renderer/form/bridge regressions covered by the focused suite.

Native verification of these two presentation/consistency fixes is still pending.

## Immediate next start point

1. Build `MASTER_TOOLBAR_SMOKE_V72.xlsm` from clean V48.
2. Native re-test first:
   - Save conflict dialog shows the full message and both **Συνέχεια / Επιστροφή** buttons,
   - patient Γιαννιά Ρομποτικό time appears as `hh:mm`, not decimal,
   - spot-check one or two other existing MASTER treatment cells for raw decimals.
3. Reconfirm conflict **Επιστροφή / Συνέχεια** behavior.
4. Verify adjacency boundary: ΦΘ `09:15-10:00` + Ανακλινόμενο `10:00-10:30` must not be treated as a conflict.
5. Then continue with edit-existing-program and optional manual names.
6. Full regression suite only after native edit smoke is clean.

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
