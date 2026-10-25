# Current Handoff

Last updated: 2026-10-25

## Read order

Start with `Documentation/ONBOARDING.md`, then `Documentation/IN_PLACE_MUTATION_ARCHITECTURE.md`, then `Documentation/DAILY_INPUT_APPLY_WORKFLOW.md`. This file is the exact resume point.

## Branch / trusted state

- Branch: `feature/outpatient-scheduling`
- Baseline `Rehab_Center_System_v27_1.xlsm` is immutable.
- Replacement-quality business logic is CLOSED end-to-end.
- Statistics/history milestone is active.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- native in-place close-day: VERIFIED; workbook stays open on MASTER
- native close-day duplicate second click: VERIFIED
- focused combined in-place gate: `57 passed in 2.65s`
- focused DAILY_INPUT/global replacement gate: `77 passed in 5.33s`
- MASTER/status-zone gate after K/L/M correction: `47 passed, 2 warnings in 3.36s`

The 2 warnings are `ZipFile.__del__` cleanup warnings, not business-logic failures. They should be cleaned later but do not change the gate result.

## Locked production interaction architecture

No normal user-facing workflow may close/reopen the workbook to complete an action.

Required model:

`validate -> compute authoritative result -> in-place native write -> semantic read-back -> Save -> continue in application`

Production actions may use temporary rollback copies while the live workbook remains open. Console windows, reopen workers and live-file replacement are forbidden.

Diagnostic-first native testing is mandatory. Do not repeat blind smokes.

## DAILY_INPUT native history

### V4

- patient absence applied in-place;
- therapist absence incorrectly routed to raw/blank `REPLACEMENTS`;
- raw-sheet routing is permanently retired.

### V5

A real ΓΙΑΝΝΙΑ cancellation exposed incomplete MASTER/workload integration. MASTER non-active states and workload counting were extended. The later focused test correctly stopped an attempted stale K-column status write because current MASTER layout is K=`ΨΥΧΟΛΟΓΟΙ`, L=`ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ`, M=`Κατάσταση`. Safety now permits only D:J therapy overlays plus M daily-status overlays; K/L remain protected.

### V6 native findings

The user entered ΓΙΑΝΝΙΑ as all-day absent/transferred (`Διακομιστήριο`). V6 showed several operational UX/data gaps:

1. MASTER incorrectly replaced a treatment cell with `ΑΠΩΝ` text. Correct behavior is to preserve recurring treatment text and apply a **red strike-through** only to sessions that actually occur on the selected date.
2. The final `Κατάσταση` column is the correct place for the dated patient reason/status.
3. `Όλη ημέρα` must cover every materialized session for that patient on the selected date. Core `DailyAbsence` logic already does this; the remaining work is presentation/persistence.
4. The initial observation that ΠΑΠΑΜΑΝΟΥ remained `6/6` is **not** treated as a confirmed defect because the user immediately noted the therapist mapping may have been mistaken. Verify actual session→therapist mapping before judging workload.
5. MASTER `Φόρτος θεραπευτών` did not navigate to `THERAPIST_DAILY`.
6. workload badges could be dragged with the mouse.
7. `THERAPIST_DAILY` lacked `Επιστροφή στο MASTER` navigation.
8. ΓΙΑΝΝΙΑ's DAILY_INPUT interruption was not recorded in the reversible daily-change ledger, so no `Επαναφορά` button existed.

## V6 remediation now on branch, awaiting focused tests

### MASTER presentation

`Python/rehab_excel/master_daily_overlay.py`

- non-active patient/cancellation/therapist states now use PRESENTATION-only treatment patches;
- recurring treatment text remains untouched;
- native writer applies red font + strike-through for `master_daily_non_active_overlay`;
- patient daily reason/status is written in the confirmed status column M when all dated sessions are non-active;
- replacement overlays still show dated effective provider/time.

### Native read-back

`Python/rehab_excel/daily_input_action_vba.py`

- source-tag presentation applies/verifies MASTER red strike and red patient-status display;
- failure still rolls back in place.

### Workload/navigation

`Python/rehab_excel/therapist_workload_strip_vba.py`

- `Φόρτος θεραπευτών` now opens `THERAPIST_DAILY`;
- fixed `Επιστροφή στο MASTER` control added;
- workload shapes are locked;
- MASTER protects drawing objects with `UserInterfaceOnly=True` so operators cannot drag badges while macros can redraw them;
- occupied-slot counting continues to ignore struck, italic replacement-source and `→` annotation lines.

`Python/rehab_excel/daily_input_change_navigation_vba.py` routes the workload button through daily-change ledger synchronization before operator restore work.

### Reversible DAILY_INPUT identity

`Python/rehab_excel/daily_input_reader.py`

Each active DAILY_INPUT row now yields a `DailyInputSourceChange` with:

- stable change id;
- exact source row;
- change type;
- subject/display name;
- reason;
- all-day/timeslot semantics.

Reader supports `ignore_rows=(...)` plus `allow_empty=True` for restore recomputation.

### Full restore projections

New:

- `Python/rehab_excel/full_daily_projection.py`
- `Python/rehab_excel/full_master_daily_projection.py`
- `Python/tools/prepare_daily_input_restore_in_place.py`

Restore preparation is read-only. It simulates removal of exactly one DAILY_INPUT source row, rebuilds the complete THERAPIST_DAILY grid from recurring/base truth, clears stale cells, preserves active replacements/surviving daily facts, resets MASTER daily presentation and reapplies surviving overlays.

### DAILY_INPUT ledger + restore transaction

New:

- `Python/rehab_excel/daily_input_change_ledger_vba.py`

It uses the existing `DAILY_CHANGE_LEDGER` visible panel. DAILY_INPUT entries carry dedicated metadata in the ledger, are rewired to `RestoreDailyInputChangeRow`, and their bulk-selection checkbox is removed so they cannot accidentally enter replacement restore.

Individual restore:

`source-row identity -> read-only restore payload -> SaveCopyAs rollback -> reset MASTER daily presentation -> clear exact DAILY_INPUT row -> apply/read-back full projection -> mark RESTORED -> refresh workload/ledger -> Save`

No close/reopen. Diagnostic:

`%TEMP%\eka_daily_input_restore_last_error.log`

### Deterministic preview installation

Both `registration_menu_vba.py` and `build_registration_menu_preview.py` now install/verify in final order:

- shell + MASTER toolbar;
- DAILY_INPUT action;
- global replacement action;
- guided double-booking review;
- workload/navigation;
- effective in-session `modDailyChangeLedger`;
- reversible DAILY_INPUT ledger extension;
- workload→ledger navigation bridge;
- close-day and registration forms.

This prevents a second operational-install pass from overwriting the new workload button routing.

## New regression coverage

Added/updated:

- `Tests/test_master_daily_overlay.py`
- `Tests/test_daily_input_source_changes.py`
- `Tests/test_daily_input_change_restore_contract.py`
- `Tests/test_therapist_navigation_and_locked_workload.py`
- `Tests/test_build_registration_menu_operational_modules.py`
- `Tests/test_production_actions_stay_in_session.py`

The production guard now includes DAILY_INPUT restore and rejects close/reopen/visible-console regressions.

## Exact resume point

Do not native-smoke V6 again.

Immediate order:

1. run the focused V7 pre-native test set covering MASTER presentation, all-day source changes, full restore plans, ledger/navigation, workload, installer and production guards;
2. debug any Python/source-contract failure before another Excel smoke;
3. when green, build fresh `DAILY_INPUT_IN_PLACE_V7.xlsm` from `DOUBLE_BOOKING_REJECTED.xlsm`;
4. native all-day ΓΙΑΝΝΙΑ transfer test:
   - recurring MASTER treatments for the selected date remain visible but red-struck;
   - M status shows the dated reason;
   - THERAPIST_DAILY affected lines are struck;
   - actual original therapist occupied slots release correctly;
   - `Φόρτος θεραπευτών` opens therapist schedule;
   - `Επιστροφή στο MASTER` exists;
   - workload badges cannot be dragged;
   - DAILY_CHANGE_LEDGER panel shows ΓΙΑΝΝΙΑ with individual `Επαναφορά`;
5. click ΓΙΑΝΝΙΑ restore and verify same workbook returns her to the daily program, refreshes MASTER/THERAPIST_DAILY/workload and marks ledger RESTORED;
6. on failure read `%TEMP%\eka_daily_input_last_error.log`, `%TEMP%\eka_replacement_last_error.log` or `%TEMP%\eka_daily_input_restore_last_error.log` before any retry;
7. after patient branch is native green, test therapist absence/global replacement continuation;
8. then continue patient/day finalized outcome history.

## Safety reminders

- baseline immutable;
- daily facts are overlays, never destructive recurring edits;
- production `.xlsm` mutation is native Excel/VBA/COM-compatible;
- no production/staging save with openpyxl;
- semantic read-back after production writes;
- Windows tests use `python -m pytest`;
- documentation updates are part of implementation.
