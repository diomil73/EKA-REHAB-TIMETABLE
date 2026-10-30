# Close-Day VBA Recovery Checkpoint

Date: 2026-10-30

## Confirmed green gate before VBA recovery

The operator reported:

`49 passed in 4.56s`

This covered read-only in-place close-day preparation together with the already-green close-day/history stack.

## Recovered VBA slice

Recovered from the later established implementation, without bulk-merging unrelated later work:

- `Python/rehab_excel/close_day_action_vba.py`
- `Tests/test_close_day_action_vba.py`

Recovered behavior:

- explicit user confirmation before closing the day,
- saves the current workbook before validation,
- invokes `prepare_close_day_in_place.py` hidden and synchronously,
- performs in-session persistence into `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY` and `DAY_HISTORY_EVENTS`,
- performs preflight duplicate/schema checks,
- performs semantic row read-back,
- rolls back partial history writes on failure,
- protects and sets technical history sheets to `xlSheetVeryHidden`,
- saves the same open workbook on success,
- does not close/reopen the workbook and does not use deferred handoff workers.

## Canonical build hook

The recovery branch intentionally does not replace the whole registration installer with the later branch version, because that later file also contains unrelated ledger/replacement changes.

Instead, `Python/tools/build_registration_menu_preview.py` now performs a narrow explicit close-day installation step after the registration preview is created:

1. pre-clean removes any stale `modCloseDayAction`,
2. a dedicated COM session calls `install_close_day_action(vbproject)`,
3. the builder verifies `modCloseDayAction` exists,
4. failure deletes the new preview output,
5. the original source hash remains protected.

## Commits

- `b7f3094` recovered `modCloseDayAction`
- `971b347` recovered initial close-day VBA tests
- `0f634d0` canonical registration build installs and verifies close-day action
- `baeae77` regression test aligned to the canonical recovery build hook

## Safety boundary

- No existing `.xlsm` was modified by this source recovery.
- No production or forensic workbook was touched.
- No replacement V2 patch/repair tooling was reintroduced.
- `modCloseDayAction` is not yet added to `REQUIRED_APP_COMPONENTS`.

## Next gate

Run the close-day VBA and canonical builder tests together with the already-green in-place/history stack. Only after that gate is green may `modCloseDayAction` be promoted into the application contract. A native Excel smoke must still use a fresh uniquely named workbook copy.
