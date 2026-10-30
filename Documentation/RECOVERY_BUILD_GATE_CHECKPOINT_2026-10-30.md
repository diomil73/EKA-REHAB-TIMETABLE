# Recovery Build Gate Checkpoint — 2026-10-30

## Purpose

This checkpoint records the first source-level recovery slice after the forensic reconstruction of the V81→V92, outpatient-recurring, and later DAILY_INPUT/operator workbook lineages.

No production workbook, forensic workbook, or existing smoke workbook was modified by this slice.

## Confirmed findings

1. The canonical `build_unified_preview.py` already contains the outpatient form stage, so the historical `frmOutpatientSchedule` / `modOutpatientSchedule` pair was not intentionally abandoned.
2. The canonical outpatient form stage installed the form and edit support but did not install `frmOutpatientScheduleManager` / `modOutpatientScheduleManager`, even though `application_contract.py` already requires those manager components.
3. The manager is a distinct required UI layer. Its edit action depends on the historical outpatient form and its registration-menu entry routes users through `ShowOutpatientScheduleManager`.
4. `build_registration_menu_preview.py` pre-cleans and reinstalls the operational stack. `install_daily_input_action()` installs global replacement and guided double-booking components.
5. Close-day is not lost from the later feature implementation. `Win32ComRegistrationMenuInstaller.install()` explicitly installs and verifies `modCloseDayAction` on the later `feature/outpatient-scheduling` branch. The recovery work branch does not yet contain that later close-day/history slice, so it must be recovered selectively rather than recreated blindly.
6. The current `REQUIRED_APP_COMPONENTS` contract remains intentionally unchanged in this first slice. It is known to be incomplete and will be expanded only after each candidate component is proven to be installed by the canonical full pipeline.

## Source changes in this slice

### `cfbffc3289214c8f7811f9550021c3b786b7d2ca`
Added regression coverage requiring the canonical outpatient form preview stage to install the outpatient schedule manager after form creation and edit support.

### `b91e57cd6e95e74a881641fd261d53c90b799bfd`
Added `install_outpatient_schedule_manager_in_workbook()` as a safe Windows/Excel wrapper around the existing manager installer. It opens only the supplied workbook copy, installs the existing manager components, saves that copy, and closes Excel without changing source workbooks.

### `ea3b982f515ebef947e9ac1855a923c7e984f108`
Updated `build_outpatient_schedule_form_preview.py` so the canonical form-preview stage now installs:

1. the outpatient scheduling form,
2. outpatient edit support,
3. the outpatient schedule manager.

This removes the previous mismatch where the application contract required manager components that the unified build stages did not explicitly create.

### `d4a83ba277344dc31783f2ceaae7367823571830`
Added unit coverage requiring final unified-workbook verification to invoke the application contract and fail closed when that contract rejects the artifact.

### `6ca4d8d1e56138d363b2ad8d66e6c6ac0d839e4a`
Updated `build_unified_preview.py` so final verification now requires:

- output exists,
- VBA project exists,
- `DAILY_INPUT` exists,
- `assert_application_contract(output)` succeeds.

An application-contract failure is converted to `UnifiedBuildError`, so an incomplete artifact is rejected and deleted by the existing fail-closed build path.

## What is deliberately NOT changed yet

- `REQUIRED_APP_COMPONENTS` has not yet been expanded.
- No close-day/history files have yet been copied from the later feature branch.
- No replacement solver or manual replacement logic was changed.
- No workbook was built or opened for mutation.
- No forensic reference workbook was touched.
- No quarantined V2 replacement patch/repair files were reused.

## Next gate

Run the focused Python tests for this slice on Windows. If green, the next source slice is:

1. prove the exact complete component set produced by the canonical unified pipeline,
2. expand `REQUIRED_APP_COMPONENTS` only with those proven components,
3. add regression coverage for missing historical outpatient form/module and other confirmed critical components,
4. selectively recover the already-implemented close-day/history files and their existing tests from the later feature branch,
5. only after Python gates are green, build a newly named recovery unified smoke copy for native Excel verification.

## Safety rule

Do not use `DAILY_INPUT_IN_PLACE_V16_OPERATOR.xlsm`, `REGISTRATION_MENU_PREVIEW.xlsm`, or any other subsystem smoke as the full application baseline. The full application must be produced by the canonical unified pipeline and must pass the strengthened application contract before native testing.
