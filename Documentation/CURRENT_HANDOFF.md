# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V78.xlsm`
- V73 scheduling/conflict behavior is native green.
- V74/V75 freeze-based pinning failed native.
- V76 command-bar sidebar produced no visible native change.
- V77 modeless sidebar did not reach the final workbook because the scheduling preview builder kept stale shell VBA.
- V78 confirms the final shell-refresh fix works: the modeless `frmMasterSidebar` appears and remains visible while the MASTER sheet scrolls.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- first MASTER toolbar gate: `39 passed in 1.71s`
- stronger freeze gate: `39 passed in 2.56s`
- command-sidebar gate: `41 passed in 2.78s`
- modeless-sidebar gate: `42 passed in 1.52s`
- final shell-refresh focused gate: `44 passed in 1.41s`

## Current focus: sidebar UI polish

Native V78 finally proves the chosen mechanism works: the modeless sidebar is visible and scroll-independent. The remaining issue is visual polish, specifically button typography/alignment.

Current polish change in `application_shell_vba.py`:

- keeps the same modeless sidebar and the same four actions,
- keeps equal button sizes and equal vertical spacing,
- uses `Segoe UI` size 10 bold,
- forces centered text with `TextAlign = 2`,
- disables wrapping so every caption stays on one line,
- uses uniform 146 x 36 buttons.

Implementation commit: `befd4418c456be5936e6ddffddd4526009a4f725`

Regression commit: `5cc195433709d2116c97d16e5c19c455956918a5`

The visual polish still needs focused pytest and a native preview check.

## Immediate next start point

1. Pull latest branch.
2. Run focused shell/sidebar tests plus `Tests/test_patient_scheduling_preview_shell_refresh.py`.
3. If green, build the next preview from clean V48.
4. Native verify the same working sidebar with improved button text alignment/spacing.
5. Once the sidebar appearance is accepted, resume edit-existing-program, optional manual names, then full regression.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal.
- MASTER is derived; `PATIENT_PLANNER` remains authoritative.
- Conflict override must be explicit; returning for correction performs no writes.
- Persistent MASTER controls must not alter or shift MASTER data columns.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Use `python -m pytest` on Windows, never bare `pytest`.
