# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V77.xlsm`
- V73 scheduling/conflict behavior is native green.
- V74/V75 freeze-based pinning failed native.
- V76 command-bar sidebar produced no visible native change.
- V77 modeless sidebar also produced no visible change because the final scheduling preview builder was not refreshing the current `modApplicationShell` at all.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- first MASTER toolbar gate: `39 passed in 1.71s`
- stronger freeze gate: `39 passed in 2.56s`
- command-sidebar gate: `41 passed in 2.78s`
- modeless-sidebar gate: `42 passed in 1.52s`
- final shell-refresh focused gate: `44 passed in 1.41s`

## Current focus: final application-shell refresh in preview builder

Native diagnostic on V77: Alt+F8 did not list `ShowMasterSidebar` at all. This proved the modeless UserForm code was not reaching the final preview workbook.

Root cause found in `Python/tools/build_patient_scheduling_slice_preview.py` / patient-scheduling build path: the final preview updated patient-program UI, renderer, conflict layout, etc., but did not reinstall the current application shell. Because V48 is the clean base, the final workbook kept its stale shell VBA.

Fix implemented:

- final preview builder now opens the completed output workbook,
- calls `install_application_shell(vbproject, workbook=workbook)`,
- verifies `modApplicationShell` exists,
- verifies `frmMasterSidebar` exists,
- verifies `Public Sub ShowMasterSidebar()` is present,
- only then saves the final preview.

Implementation commit: `9d3296a0cb8600b77a8cf1b6187494993a711cb7`

Regression test added: `Tests/test_patient_scheduling_preview_shell_refresh.py`

Test commit: `e5851cc3b0b1f1db21679a70bbae37994de5d704`

Focused gate after the build-pipeline fix is green: `44 passed in 1.41s`.

## Immediate next start point

1. Build `MASTER_TOOLBAR_SMOKE_V78.xlsm` from clean V48.
2. Open V78 and use Alt+F8 to confirm `ShowMasterSidebar` now exists.
3. If it exists, run it manually once if needed and verify the modeless sidebar appears.
4. Then verify scroll independence, all four actions, hide on other sheet, restore on return.
5. Once sidebar UX is native green, resume edit-existing-program, optional manual names, then full regression.

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
