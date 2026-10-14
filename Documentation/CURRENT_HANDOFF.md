# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V73.xlsm`
- V73 builds and runs correctly in native Excel.
- Next preview after the new MASTER toolbar pinning gate: `MASTER_TOOLBAR_SMOKE_V74.xlsm` from clean V48.

## What is verified in native Excel

- The patient-centric multi-treatment form opens and is usable.
- First patient-program save succeeds.
- ΦΘ self-conflict false alarm is fixed.
- Conflict validation before Auto and Save works.
- Save conflict dialog shows the full message and both **Συνέχεια / Επιστροφή** buttons.
- Manual ΦΘ occupied-slot diagnostics identify the real blocker with day/time/provider/patient context.
- `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15` is accepted correctly as exact adjacency.
- True treatment overlaps produce detailed treatment/day/time diagnostics.
- The previously observed Γιαννιά Ρομποτικό decimal display issue is resolved in normal use after re-entry/save; the user reports that everything plays correctly.

## V73 native result

The V73 scheduling/conflict batch is considered native green:

- precise manual ΦΘ occupancy diagnostic: passed,
- exact adjacency: passed,
- true overlap diagnostics: passed,
- conflict dialog runtime layout: passed,
- normal Ρομποτικό time rendering after save/edit: passed.

Latest user-run source/integration gates:

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`

## Current focus: persistent MASTER toolbar

The user requested that the four MASTER action buttons remain visible while scrolling, or alternatively become a left sidebar.

Chosen implementation: keep the existing top toolbar and freeze the first three MASTER rows. This is safer than a sidebar because it does not insert or shift any MASTER data columns.

Source implementation in `application_shell_vba.py`:

- new `PinMasterToolbar()` applies `SplitRow = 3` and `FreezePanes = True`,
- horizontal split remains zero,
- the pin is reapplied on application entry and whenever the shell returns to/activates `MASTER_SCHEDULE`,
- row 2 therefore remains visible while the patient list scrolls vertically.

Implementation commit:

`346247ea1dde7f9a3529bf6c25d786b08409699b`

Regression test added:

`Tests/test_master_toolbar_pinned_view.py`

Test commit:

`0b5eb5a18a2ba0307bf57d61099d38b9d6a612fb`

This toolbar change is not yet user-run verified. Do not mark it green until the focused pytest and native V74 smoke pass.

## Immediate next start point

1. Run focused toolbar/application-shell tests.
2. If green, build `MASTER_TOOLBAR_SMOKE_V74.xlsm` from clean V48.
3. Native verify on MASTER:
   - scroll well below the first screen and confirm the four row-2 buttons remain visible,
   - verify all four buttons still work,
   - go to DAILY_INPUT / THERAPIST_DAILY and return to MASTER, confirming the toolbar is pinned again.
4. After that resume the patient-program milestone:
   - native edit-existing-program flow end to end,
   - optional manual names persist/reload,
   - full regression suite.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal.
- MASTER is a derived view; `PATIENT_PLANNER` remains authoritative.
- Conflict override must be explicit; returning for correction performs no writes.
- MASTER toolbar pinning must not alter or shift MASTER data columns.

## Recovery order if context is lost

Read these in order:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests for the current milestone

Use `python -m pytest` on Windows, never bare `pytest`.
