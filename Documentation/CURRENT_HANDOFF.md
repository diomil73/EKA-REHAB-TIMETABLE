# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V75.xlsm`
- V73 scheduling/conflict behavior is native green.
- V74 and V75 navigation actions work, but neither freeze-based implementation keeps the four MASTER buttons visible during vertical scrolling.
- Freeze-based pinning is now abandoned.
- Next preview will be `MASTER_TOOLBAR_SMOKE_V76.xlsm` from clean V48 and will test a scroll-independent left MASTER sidebar.

## What is verified in native Excel

- The patient-centric multi-treatment form opens and is usable.
- First patient-program save succeeds.
- ΦΘ self-conflict false alarm is fixed.
- Conflict validation before Auto and Save works.
- Save conflict dialog shows the full message and both **Συνέχεια / Επιστροφή** buttons.
- Manual ΦΘ occupied-slot diagnostics identify the real blocker with day/time/provider/patient context.
- `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15` is accepted correctly as exact adjacency.
- True treatment overlaps produce detailed treatment/day/time diagnostics.
- The previously observed Γιαννιά Ρομποτικό decimal display issue is resolved in normal use after re-entry/save.
- V74/V75 confirm the toolbar actions/navigation still work; only persistent visibility failed.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- first focused MASTER toolbar/application-shell gate: `39 passed in 1.71s`
- focused gate after selection-based freeze fix: `39 passed in 2.56s`
- focused sidebar/application-shell gate: `41 passed in 2.78s`

## Current focus: scroll-independent MASTER sidebar

Two native freeze-pane attempts failed to keep worksheet shape-buttons continuously visible in V74 and V75. The project has therefore pivoted to a different mechanism rather than continuing to tune `FreezePanes`.

New implementation in `application_shell_vba.py`:

- creates a temporary Excel command bar named `EKA_MASTER_SIDEBAR`, docked on the left,
- contains the same four actions: **Διαχείριση Μητρώων**, **Απουσίες / Ακυρώσεις**, **Ημερήσιο πρόγραμμα**, **Save & Exit**,
- is shown only while `MASTER_SCHEDULE` is active,
- is hidden on other sheets,
- hides the legacy `ekaToolbar_` worksheet shapes while the sidebar is active to avoid duplicate controls,
- is deleted when the application shell exits.

Implementation commit:

`ad7bfbd7045a445eddb747d5b5f91d482c6046b0`

Regression commits:

- `37f38b41df2f9a3d18cfba04046b7c45bbefcf4d`
- `5fcbc3159a7c603a556be788a5596fd69fd34e34`

Focused source gate is green: `41 passed in 2.78s`.

## Immediate next start point

1. Build `MASTER_TOOLBAR_SMOKE_V76.xlsm` from clean V48.
2. Native verify:
   - left sidebar is visible on MASTER,
   - vertical scroll does not move it,
   - all four actions work,
   - navigating away hides it,
   - returning to MASTER restores it.
3. Once sidebar UX is native green, resume:
   - edit-existing-program end to end,
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
- Persistent MASTER controls must not alter or shift MASTER data columns.

## Recovery order if context is lost

Read these in order:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests for the current milestone

Use `python -m pytest` on Windows, never bare `pytest`.
