# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V73.xlsm`
- V73 builds and opens successfully.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use.
- Conflict validation before Auto and Save works.
- Save conflict dialog now shows the full message and both **Συνέχεια / Επιστροφή** buttons.
- Manual ΦΘ occupied-slot diagnostics now identify the real blocker with concrete day/time/provider/patient context.
- `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15` is accepted correctly as exact adjacency, not overlap.
- True treatment overlaps produce detailed treatment/day/time diagnostics.

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

The V73 native smoke passed all intended behavioral checks except one remaining display issue:

- precise manual ΦΘ occupancy diagnostic: passed,
- exact adjacency Ανακλινόμενο 11:00-11:30 + ΦΘ 11:30-12:15: passed,
- true overlap diagnostics: passed,
- conflict dialog runtime layout: passed,
- existing Ρομποτικό time for patient Γιαννιά is still displayed as a raw decimal in MASTER.

The user is now testing whether re-entering/saving that same Ρομποτικό time changes the cell to `hh:mm`.

Interpretation of that test:

- if re-save changes it to `hh:mm`, this points to stale legacy MASTER content that the initial all-patient refresh did not fully normalize,
- if it remains decimal after re-save, there is still another write/render path that writes the raw Excel time serial after the composite renderer.

## Latest confirmed test checkpoint

User-run gates after the latest scheduling changes:

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`

Source/integration verification is green, and V73 native behavior is green except for the single Γιαννιά decimal-time display issue.

## Immediate next start point

1. Wait for the user's result after re-entering/saving Γιαννιά Ρομποτικό time.
2. If it becomes `hh:mm`, investigate stale existing MASTER content / initial refresh normalization.
3. If it stays decimal, trace the post-save write paths that can overwrite the composite renderer output.
4. Once the decimal-time issue is closed, continue with edit-existing-program and optional manual names.
5. Full regression suite only after native edit smoke is clean.

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
