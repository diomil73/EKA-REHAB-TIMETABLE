# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V73.xlsm`
- V73 builds and runs correctly in native Excel.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use.
- Conflict validation before Auto and Save works.
- Save conflict dialog shows the full message and both **Συνέχεια / Επιστροφή** buttons.
- Manual ΦΘ occupied-slot diagnostics identify the real blocker with concrete day/time/provider/patient context.
- `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15` is accepted correctly as exact adjacency, not overlap.
- True treatment overlaps produce detailed treatment/day/time diagnostics.
- The previously observed Γιαννιά Ρομποτικό decimal display issue is now confirmed resolved in normal use after re-entry/save; the user reports that everything plays correctly.

## Conflict-validation behavior

A common `validate-vba` bridge command checks pairwise patient-program overlaps using the duration contract:

- standard treatments: 45 minutes,
- Ανακλινόμενο: 30 minutes,
- only shared weekdays count,
- adjacent intervals are allowed.

The patient-program UI:

- validates before **Αυτόματη συμπλήρωση**,
- validates before any atomic Save write begins,
- shows **Συνέχεια** and **Επιστροφή** on Save conflicts,
- **Επιστροφή** cancels before writes,
- **Συνέχεια** is an explicit override.

## V73 native result

The V73 native smoke is now considered green for the current scheduling batch:

- precise manual ΦΘ occupancy diagnostic: passed,
- exact adjacency Ανακλινόμενο 11:00-11:30 + ΦΘ 11:30-12:15: passed,
- true overlap diagnostics: passed,
- conflict dialog runtime layout: passed,
- Ρομποτικό time rendering in normal save/edit use: passed.

## Latest confirmed test checkpoint

User-run gates after the latest scheduling changes:

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`

Source/integration verification and the corresponding V73 native behavior are green.

## Immediate next start point

1. Native-test the existing-patient edit flow end to end:
   - Basic menu → **Επεξεργασία ασθενή**,
   - save registry changes,
   - same patient-program form opens for the same PatientID,
   - existing program is preloaded from `PATIENT_PLANNER`,
   - modify one or more treatments,
   - final atomic save succeeds,
   - MASTER reflects the edited composite program correctly.
2. Then verify optional manual names persist/reload for Ρομποτικό / Εργο / Λογο / Ψυχ.
3. Run final duration/recliner smoke if needed.
4. Run the full regression suite only after native edit smoke is clean.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual day/time/provider constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration currently used by the solver: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal and is not an overlap.
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
