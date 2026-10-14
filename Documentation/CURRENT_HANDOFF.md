# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V74.xlsm`
- V73 scheduling/conflict behavior is native green.
- V74 navigation buttons work, but the four MASTER buttons do **not** remain visible while vertically scrolling.
- Next preview after the stronger pinning fix: `MASTER_TOOLBAR_SMOKE_V75.xlsm` from clean V48.

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
- V74 confirms the toolbar actions and navigation still work.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- focused MASTER toolbar/application-shell gate: `39 passed in 1.71s`

## Current focus: persistent MASTER toolbar

The user wants the four MASTER action buttons continuously visible while scrolling.

The first runtime approach used `SplitRow = 3` + `FreezePanes = True`. V74 showed that this did not actually keep the buttons visible in native Excel, even though navigation itself remained correct.

The stronger fix now uses Excel's canonical selection-based freeze procedure inside `PinMasterToolbar()`:

- activate/confirm `MASTER_SCHEDULE`,
- clear existing freeze and splits,
- select `A4`,
- set `FreezePanes = True`,
- restore selection to `A1`.

This should freeze rows 1–3, keeping row 2 and its four buttons visible during vertical scrolling. Runtime reapplication on entry/return remains in place.

Implementation commit:

`3138f6b792677e1183af23c6a5d467b41711d003`

Regression update:

`71a50dc34bf9fc5aef3fb1c74b11812d929a81a8`

This stronger pinning fix is not yet user-run verified.

## Immediate next start point

1. Run the focused toolbar/application-shell tests again.
2. If green, build `MASTER_TOOLBAR_SMOKE_V75.xlsm` from clean V48.
3. Native verify:
   - scroll far down on MASTER and confirm all four row-2 buttons remain visible,
   - verify all four buttons still work,
   - navigate away and back to MASTER and confirm pinning remains/reapplies.
4. Once toolbar pinning is native green, resume:
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
- MASTER toolbar pinning must not alter or shift MASTER data columns.

## Recovery order if context is lost

Read these in order:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests for the current milestone

Use `python -m pytest` on Windows, never bare `pytest`.
