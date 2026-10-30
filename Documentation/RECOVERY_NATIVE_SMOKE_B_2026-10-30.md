# Recovery Native Smoke B - 2026-10-30

## Artifact boundary

Native smoke artifact:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm`

This workbook is evidence only after the native smoke. Do not overwrite, repair in place, or reuse it as the next experimentation surface.

Verified before opening:

- source `Rehab_Center_System_v27_1.xlsm` SHA remained unchanged,
- `DAILY_INPUT` exists,
- `DAILY_INPUT!B2` is `30/10/2026`,
- smoke artifact size was 382507 bytes.

## Native failure 1: outpatient schedule entry point

From the central menu, the outpatient schedule action reported that the form was not available for the workbook.

Source inspection showed that the current menu handler catches every error raised by `frmOutpatientSchedule.Show` and translates it to the same `not installed` message. Therefore the native observation does not prove that `frmOutpatientSchedule` is missing. It may represent an initialization, compile, reference, or runtime error inside the form.

The outpatient form source on `work/history-statistics-rebuild` was compared with the later established feature source and is the same lineage. The next fresh smoke must expose the original `Err.Number` and `Err.Description` instead of masking it as a missing-form condition.

Do not patch smoke B.

## Native failure 2: DAILY_INPUT close/reopen path

The operator entered a patient absence. Excel launched a black console/process, later closed it, prompted to continue, and ultimately reported a workbook-close failure (`Close of object Workbook failed`).

Source inspection confirmed that the recovery branch still contained the older DAILY_INPUT authoritative handoff that launches `authoritative_commit_worker_cli.py` and executes `ThisWorkbook.Close`.

This is incompatible with the established in-place mutation architecture:

`validate -> compute -> in-place native write -> semantic read-back -> rollback on failure -> save same workbook`.

The later established DAILY_INPUT implementation instead uses a read-only Python preparation step, validated TSV patches, native in-session application, semantic read-back, rollback copy, operational projection refresh and same-workbook save. It does not close/reopen the user workbook.

## Recovery work started after smoke B

Recovered as read-only prerequisites, before changing installed VBA:

- `Python/tools/prepare_daily_input_in_place.py`
- `Python/tools/prepare_global_replacements_in_place.py`
- `Tests/test_prepare_daily_input_in_place.py`
- `Tests/test_prepare_global_replacements_in_place.py`

The global preparer uses the established authoritative global solver and explicit guided double-booking review. It does not contain or depend on the quarantined manual replacement V2 patch/normalizer tools.

## Safety gate before smoke C

1. The two read-only preparer test suites must be green.
2. Replace the old DAILY_INPUT close/reopen VBA action only after its full in-place dependency chain is recovered and tested.
3. Add an outpatient-menu diagnostic that preserves the real native error number/description.
4. Run source regression tests including an explicit prohibition on `ThisWorkbook.Close`, `Application.OnTime`, and authoritative reopen/worker plumbing in production actions.
5. Build a fresh uniquely named `UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm` only after those gates are green.
6. Do not run `CloseOperationalDay` until outpatient opening and DAILY_INPUT native smoke are both clean on the fresh copy.

## Exclusions

Do not reintroduce:

- `manual_replacement_operator_v2_patch.py`,
- `manual_operator_v2_source_normalizer.py`,
- `repair_manual_replacement_operator_v2.py`,
- any broad merge/cherry-pick from the later feature branch.
