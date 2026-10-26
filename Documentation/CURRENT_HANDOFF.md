# Current Handoff

Last updated: 2026-10-26

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
- V7 pre-native focused gate: `63 passed, 6 warnings in 5.92s`

The current warnings are `PytestUnraisableExceptionWarning` from Python 3.14/openpyxl `ZipFile.__del__` cleanup in disposable test workbooks. They are not business-logic failures. Keep them as technical cleanup debt.

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

## V6 remediation now on branch

### MASTER presentation

`Python/rehab_excel/master_daily_overlay.py`

- non-active patient/cancellation/therapist states use PRESENTATION-only treatment patches;
- recurring treatment text remains untouched;
- native writer applies red font + strike-through for `master_daily_non_active_overlay`;
- patient daily reason/status is written in confirmed status column M when all dated sessions are non-active;
- replacement overlays still show dated effective provider/time.

### Native read-back

`Python/rehab_excel/daily_input_action_vba.py`

- source-tag presentation applies/verifies MASTER red strike and red patient-status display;
- failure still rolls back in place.

### Workload/navigation

`Python/rehab_excel/therapist_workload_strip_vba.py`

- `Φόρτος θεραπευτών` opens `THERAPIST_DAILY`;
- fixed `Επιστροφή στο MASTER` control added;
- workload shapes are locked;
- MASTER protects drawing objects with `UserInterfaceOnly=True` so operators cannot drag badges while macros can redraw them;
- occupied-slot counting ignores struck, italic replacement-source and `→` annotation lines.

`Python/rehab_excel/daily_input_change_navigation_vba.py` routes the workload button through daily-change ledger synchronization before operator restore work.

### Reversible DAILY_INPUT identity

`Python/rehab_excel/daily_input_reader.py`

Each active DAILY_INPUT row yields a `DailyInputSourceChange` with stable change id, exact source row, change type, subject/display name, reason and all-day/timeslot semantics.

Reader supports `ignore_rows=(...)` plus `allow_empty=True` for restore recomputation.

### Full restore projections

New:

- `Python/rehab_excel/full_daily_projection.py`
- `Python/rehab_excel/full_master_daily_projection.py`
- `Python/tools/prepare_daily_input_restore_in_place.py`

Restore preparation is read-only. It simulates removal of exactly one DAILY_INPUT source row, rebuilds the complete THERAPIST_DAILY grid from recurring/base truth, clears stale cells, preserves active replacements/surviving daily facts, resets MASTER daily presentation and reapplies surviving overlays.

### DAILY_INPUT ledger + restore transaction

`Python/rehab_excel/daily_input_change_ledger_vba.py` uses the existing `DAILY_CHANGE_LEDGER` visible panel. DAILY_INPUT entries carry dedicated metadata, are rewired to `RestoreDailyInputChangeRow`, and their bulk-selection checkbox is removed so they cannot accidentally enter replacement restore.

Individual restore:

`source-row identity -> read-only restore payload -> SaveCopyAs rollback -> reset MASTER daily presentation -> clear exact DAILY_INPUT row -> apply/read-back full projection -> mark RESTORED -> refresh workload/ledger -> Save`

No close/reopen. Diagnostic:

`%TEMP%\eka_daily_input_restore_last_error.log`

## V7 build blockers fixed

### Guided double-booking UserForm recreation

The second preview pass tried to remove/recreate `frmDoubleBookingReview` even though the first pass had already installed it. Excel/VBE retained the UserForm Designer/name and raised COM error 1004.

Fix: second pass is idempotent. Existing module+form are kept and only missing/partial installs are recreated.

Relevant commits:

- `3052b371354f740c0e41ad1f85131b5e89b7fc04`
- `088f833e831fe937b898928d0b88a82e0801d539`

### Workload RowHeight build failure

Build then progressed to `installing therapist workload strip` and failed on `master.Rows(3).RowHeight = 28`. RowHeight is cosmetic only and is now best-effort in both Python installer and runtime VBA.

Relevant commits:

- `d9a9572dcb5a373812cec8e599bd9dc4371ee2d0`
- `62fa398f33c409accbc714f62f9e47b01cc05e84`

## V7 first runtime diagnostic

After the builder blockers were removed, the first actual DAILY_INPUT runtime attempt failed with operator message that the daily state was not updated.

Diagnostic captured:

```text
timestamp=2026-10-26 00:13:25
phase=daily input action
patch_index=0
sheet=
cell=
error_number=1004
error=Το κελί ή το γράφημα που προσπαθείτε να αλλάξετε βρίσκεται σε προστατευμένο φύλλο...
```

Interpretation:

- failure occurred before any validated PATCH (`patch_index=0`);
- this is not a ΓΙΑΝΝΙΑ business-rule failure and not a patch-target/read-back failure;
- the outer action had only generic phase reporting, so the exact pre-payload step was hidden;
- no `Workbook_BeforeSave` implementation exists in repo source, therefore legacy/binary workbook event behavior remains a plausible source of the protected-sheet mutation during the initial `ThisWorkbook.Save`.

Fix now on branch:

- outer DAILY_INPUT action captures explicit phases: `preflight`, `resolve preparation tool`, `initial save`, `prepare payload paths`, `python preparation`, `apply payload`, `refresh operational views`, replacement continuation;
- the initial `ThisWorkbook.Save` runs with `Application.EnableEvents = False` and always restores the prior event state;
- Python preparation failure now writes the hidden diagnostic response instead of only showing a generic operator message;
- operator messages remain non-technical;
- cleanup restores the original event setting.

Relevant commits:

- `d40d0f18d5506831633035c9592da33817423fd9` event-safe initial save + phase-aware diagnostics
- `42ab9a1dd2c71fa902b6a4e69d9c865410670744` regression guards for initial-save event isolation and preparation diagnostics

## New regression coverage

Added/updated:

- `Tests/test_master_daily_overlay.py`
- `Tests/test_daily_input_source_changes.py`
- `Tests/test_daily_input_change_restore_contract.py`
- `Tests/test_therapist_navigation_and_locked_workload.py`
- `Tests/test_build_registration_menu_operational_modules.py`
- `Tests/test_daily_input_action_vba.py`
- `Tests/test_production_actions_stay_in_session.py`

The production guard includes DAILY_INPUT restore and rejects close/reopen/visible-console regressions.

## Exact resume point

Do not reuse an older V7 workbook because installed VBA must contain the latest event-safe DAILY_INPUT action.

Immediate order:

1. pull latest branch;
2. run focused DAILY_INPUT action + production/build guards;
3. rebuild a fresh `DAILY_INPUT_IN_PLACE_V8.xlsm` from `Excel/previews/DOUBLE_BOOKING_REJECTED.xlsm`;
4. repeat only the ΓΙΑΝΝΙΑ all-day `Διακομιστήριο` patient branch;
5. if it fails, immediately read `%TEMP%\eka_daily_input_last_error.log`; the new phase must identify the exact step and no blind retry is allowed;
6. if application succeeds, verify MASTER red strike, M status, THERAPIST_DAILY affected lines, actual therapist occupied-slot changes, workload navigation/locking and individual ΓΙΑΝΝΙΑ `Επαναφορά`;
7. click ΓΙΑΝΝΙΑ restore and verify same workbook stays open, returns her to daily program, refreshes MASTER/THERAPIST_DAILY/workload, clears exact DAILY_INPUT fact and marks ledger row RESTORED;
8. after patient branch is native green, test therapist absence/global replacement continuation;
9. then continue patient/day finalized outcome history;
10. later clean Python 3.14/openpyxl ZipFile destructor warnings without weakening workbook-safety tests.

## Safety reminders

- baseline immutable;
- daily facts are overlays, never destructive recurring edits;
- production `.xlsm` mutation is native Excel/VBA/COM-compatible;
- no production/staging save with openpyxl;
- semantic read-back after production writes;
- Windows tests use `python -m pytest`;
- documentation updates are part of implementation.
