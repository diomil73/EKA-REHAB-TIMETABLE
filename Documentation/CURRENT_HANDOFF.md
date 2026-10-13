# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested successfully by the user: `MASTER_TOOLBAR_SMOKE_V70.xlsm`

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- MASTER patient-program rendering works after save instead of showing raw decimals or literal `hh:mm`.
- MASTER rendering uses composite text from `PATIENT_PLANNER`: time + days + provider/name where applicable.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use: changing a non-ΦΘ treatment and pressing **Αυτόματη συμπλήρωση** no longer falsely reports unchanged ΦΘ.

## New conflict-validation issue found in V70

Native smoke exposed a more important consistency gap:

- a manual Λογο time was placed on top of ΦΘ on a shared day,
- **Αυτόματη συμπλήρωση** detected that something was wrong, but the message was too vague,
- **Αποθήκευση προγράμματος** accepted and persisted the conflict,
- MASTER then displayed the conflicting program normally.

This means automatic completion and final save did not share the same conflict gate.

## Fix implemented and focused-test verified

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

## Latest confirmed test checkpoint

Latest user-reported focused run after the conflict-validation batch:

`59 passed in 5.89s`

This closes the focused regression gate for `validate-vba`, detailed diagnostics, the conflict dialog, and the surrounding patient-program VBA paths. Native Excel verification is now the next gate.

## Immediate next start point

1. Build a fresh preview from clean V48 as `MASTER_TOOLBAR_SMOKE_V71.xlsm`.
2. Native smoke the conflict scenario:
   - ΦΘ and Λογο on the same weekday with overlapping intervals,
   - **Αυτόματη συμπλήρωση** must name the exact day and both treatments/times,
   - **Αποθήκευση** must show **Συνέχεια / Επιστροφή** before any write,
   - choose **Επιστροφή** and verify nothing changed in MASTER,
   - repeat and choose **Συνέχεια** and verify intentional override saves.
3. Verify adjacency boundary: ΦΘ `09:15-10:00` + Ανακλινόμενο `10:00-10:30` must not be treated as a conflict.
4. Then continue with the still-unverified edit workflow:
   - **Βασικό μενού → Επεξεργασία ασθενή**,
   - save registry changes,
   - same patient-program form opens,
   - existing program is preloaded,
   - modify and atomic-save it.
5. Verify optional manual names for Ρομποτικό / Εργο / Λογο / Ψυχ persist and reload.
6. Run the full regression suite only after the native edit smoke is clean.

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
