# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested successfully by the user: `MASTER_TOOLBAR_SMOKE_V70.xlsm`
- `MASTER_TOOLBAR_SMOKE_V71.xlsm` build attempt stopped safely during conflict-dialog installation with a VBE COM sizing error.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- MASTER patient-program rendering works after save instead of showing raw decimals or literal `hh:mm`.
- MASTER rendering uses composite text from `PATIENT_PLANNER`: time + days + provider/name where applicable.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use: changing a non-ΦΘ treatment and pressing **Αυτόματη συμπλήρωση** no longer falsely reports unchanged ΦΘ.

## Conflict-validation issue found in V70

Native smoke exposed a more important consistency gap:

- a manual Λογο time was placed on top of ΦΘ on a shared day,
- **Αυτόματη συμπλήρωση** detected that something was wrong, but the message was too vague,
- **Αποθήκευση προγράμματος** accepted and persisted the conflict,
- MASTER then displayed the conflicting program normally.

This means automatic completion and final save did not share the same conflict gate.

## Conflict-validation fix implemented and focused-test verified

A common `validate-vba` bridge command now checks pairwise patient-program overlaps using the same duration contract:

- standard treatments: 45 minutes,
- Ανακλινόμενο: 30 minutes,
- only shared weekdays count,
- adjacent intervals are allowed.

Detailed diagnostics now have the form:

`Δευτέρα: ΦΘ 09:00-09:45 συγκρούεται με Λογο 09:00-09:45.`

The patient-program UI is patched so that:

- **Αυτόματη συμπλήρωση** runs detailed conflict validation first and stops with the explicit day/treatment/time message,
- **Αποθήκευση** runs the same validation before any atomic write begins,
- if a conflict exists, a modal form shows two explicit choices: **Συνέχεια** and **Επιστροφή**,
- **Επιστροφή** cancels the save before anything is written,
- **Συνέχεια** is an explicit user override and allows the conflicting program to be saved intentionally.

## V71 installer blocker and hardening

The first V71 build stopped with:

`SAFETY STOP: Patient scheduling UI installation failed during installing MASTER patient-program renderer: Property '<unknown>.Width' can not be set.`

The cause was isolated to the new `frmPatientProgramConflict` installer. Unlike the older form installers, it wrote `designer.Width/Height` and control `Width/Height` directly, without guarding against VBE COM property rejection.

Fix implemented:

- conflict-form designer Width/Height are now best-effort inside a guarded block,
- conflict controls set Left/Top in a guarded block,
- conflict controls set Width/Height in a separate guarded block,
- a regression test now simulates a VBE control that raises on Width/Height and requires the installer helper to continue without failing.

This sizing hardening is implemented but not yet user-test-verified.

## Latest confirmed test checkpoint

Latest user-reported focused run before the V71 sizing hardening:

`59 passed in 5.89s`

The new VBE Width/Height tolerance regression must be rerun before rebuilding V71.

## Immediate next start point

1. Pull latest `feature/outpatient-scheduling`.
2. Run focused tests including `Tests/test_patient_program_master_render_vba.py` and the existing conflict-validation gate.
3. If green, rebuild `MASTER_TOOLBAR_SMOKE_V71.xlsm` from clean V48.
4. Native smoke the conflict scenario:
   - ΦΘ and Λογο on the same weekday with overlapping intervals,
   - **Αυτόματη συμπλήρωση** must name the exact day and both treatments/times,
   - **Αποθήκευση** must show **Συνέχεια / Επιστροφή** before any write,
   - choose **Επιστροφή** and verify nothing changed in MASTER,
   - repeat and choose **Συνέχεια** and verify intentional override saves.
5. Verify adjacency boundary: ΦΘ `09:15-10:00` + Ανακλινόμενο `10:00-10:30` must not be treated as a conflict.
6. Then continue with the still-unverified edit workflow and optional manual names.
7. Run the full regression suite only after the native edit smoke is clean.

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
