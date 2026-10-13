# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested successfully by the user: `MASTER_TOOLBAR_SMOKE_V70.xlsm`
- `MASTER_TOOLBAR_SMOKE_V71.xlsm` has not yet built successfully; two installer blockers were found and hardened in sequence.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- MASTER patient-program rendering works after save instead of showing raw decimals or literal `hh:mm`.
- MASTER rendering uses composite text from `PATIENT_PLANNER`: time + days + provider/name where applicable.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use: changing a non-ΦΘ treatment and pressing **Αυτόματη συμπλήρωση** no longer falsely reports unchanged ΦΘ.

## Conflict-validation issue found in V70

Native smoke exposed a consistency gap:

- a manual Λογο time was placed on top of ΦΘ on a shared day,
- **Αυτόματη συμπλήρωση** detected that something was wrong, but the message was too vague,
- **Αποθήκευση προγράμματος** accepted and persisted the conflict,
- MASTER then displayed the conflicting program normally.

The new common conflict gate is intended to fix this for both Auto and Save.

## Conflict-validation fix implemented

A common `validate-vba` bridge command checks pairwise patient-program overlaps using the duration contract:

- standard treatments: 45 minutes,
- Ανακλινόμενο: 30 minutes,
- only shared weekdays count,
- adjacent intervals are allowed.

Detailed diagnostics have the form:

`Δευτέρα: ΦΘ 09:00-09:45 συγκρούεται με Λογο 09:00-09:45.`

The patient-program UI is patched so that:

- **Αυτόματη συμπλήρωση** validates first and stops with explicit day/treatment/time diagnostics,
- **Αποθήκευση** validates before any atomic write begins,
- conflicts show a modal form with **Συνέχεια** and **Επιστροφή**,
- **Επιστροφή** cancels before any write,
- **Συνέχεια** is an explicit override.

## V71 installer blockers and hardening

### 1. Conflict form COM sizing

First V71 build stopped with:

`Property '<unknown>.Width' can not be set.`

Cause: the new conflict form directly wrote designer/control Width/Height. Fix: sizing/positioning is now best-effort with COM guards, matching the proven pattern used by older form installers. A regression test simulates Width/Height rejection.

### 2. PatientProgramRuntime validation marker

Next V71 build progressed past conflict-form installation but stopped with:

`MASTER render patch marker missing: patient-program validation runtime`

Cause: `_patch_runtime_validation()` required one exact multi-function text block containing `CompletePatientProgram` immediately followed by `FindPatientDisplayName`. Native VBE code differed enough that the exact marker did not match.

Fix implemented:

- runtime validation patch is now procedure-scoped,
- it locates only `Public Function CompletePatientProgram(...)`,
- finds that procedure's `End Function`,
- inserts `ValidatePatientProgram(...)` immediately after it,
- does not depend on which function follows next or on internal spacing,
- remains CRLF-safe and idempotent,
- gives precise diagnostics if `CompletePatientProgram` or its `End Function` is genuinely missing.

Regression coverage includes altered procedure body/spacing and idempotence.

## Latest confirmed test checkpoint

Latest user-reported focused gate after both V71 installer hardening changes:

`62 passed in 3.34s`

This confirms the conflict-form COM tolerance, procedure-scoped runtime validation patcher, bridge diagnostics, patient-program form patches, and related automatic scheduling paths are green at source-test level.

## Immediate next start point

1. Rebuild `MASTER_TOOLBAR_SMOKE_V71.xlsm` from clean V48.
2. If installer succeeds, native smoke the conflict scenario:
   - ΦΘ and Λογο on the same weekday with overlapping intervals,
   - **Αυτόματη συμπλήρωση** must name the exact day and both treatments/times,
   - **Αποθήκευση** must show **Συνέχεια / Επιστροφή** before any write,
   - choose **Επιστροφή** and verify nothing changed in MASTER,
   - repeat and choose **Συνέχεια** and verify intentional override saves.
3. Verify adjacency boundary: ΦΘ `09:15-10:00` + Ανακλινόμενο `10:00-10:30` must not be treated as a conflict.
4. Then continue with edit-existing-program and optional manual names.
5. Full regression suite only after native edit smoke is clean.

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
