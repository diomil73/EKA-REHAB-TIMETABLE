# Recovery handoff J → M — 2026-10-31

## Scope

This checkpoint records the full recovery sequence from the first native J restore attempt through the current pre-M gate. It is intended as the authoritative handoff for the next session so the same debugging tracks are not repeated.

## Safety invariants preserved

- Production workbook `Rehab_Center_System_v27_1.xlsm` has remained unchanged throughout these attempts.
- Every native candidate uses a fresh filename. Failed or partially validated artifacts are frozen as evidence and are never repaired in place.
- No close/reopen production architecture was reintroduced.
- No `Application.OnTime` restore transport was reintroduced.
- Therapist-absence restore remains disabled until related replacement facts can be rolled back atomically.
- `CloseOperationalDay` remains blocked until patient/session restore is native-proven.

## J build and findings

`UNIFIED_CLOSE_DAY_RECOVERY_20261031_J.xlsm` built successfully with application contract 24/24.

Native patient-absence smoke exposed two defects:

1. MASTER status/reason remained struck through even though the treatment strike was expected.
2. Per-row restore failed during preparation with `At least one non-active daily state is required`.

J was frozen as evidence.

### J formatting diagnosis and fix

The status overlay already set `Strikethrough=False` at patch level, but a later `ApplyMasterView` could reintroduce formatting.

A post-view normalization helper was added so the MASTER status column is normalized after `ApplyMasterView`.

Relevant commits:
- `d828aff` post-`ApplyMasterView` status normalization
- `1a43288` regression
- `ad43854` J findings checkpoint

### J restore zero-state diagnosis and fix

Read-only diagnostic on frozen J identified ACTIVE change:

`daily-input:2026-10-30:patient:21`

Before the fix the preparer failed with:

`SAFETY STOP: At least one non-active daily state is required`

Root cause: restoring the only active daily change returns the day to a fully ACTIVE state, but the generic patient-centric preview builder intentionally requires at least one non-active state.

The generic preview invariant was not weakened. A bounded zero-change restore path was added inside `prepare_daily_change_restore_in_place.py`.

When the restored day is fully active, the restore preparer now rebuilds only the affected cells from recurring/base truth:
- `THERAPIST_DAILY` original slot,
- MASTER treatment cell,
- MASTER patient status cell.

Relevant commits:
- `5129b73` clean-day restore path
- `e1695cb` zero-state restore regressions
- `e6f9572` zero-state restore checkpoint

Read-only verification on the same frozen J then succeeded:

- `DAILY CHANGE RESTORE IN-PLACE PAYLOAD VERIFIED`
- Date: `2026-10-30`
- Change count: `1`
- Source row: `21`
- Patches: `3`
- J hash unchanged: `True`

Payload contained:
- `SOURCE_CLEAR DAILY_INPUT 21 A E`
- exact `RESTORE_CHANGE` id
- `THERAPIST_DAILY` base-slot restoration
- MASTER treatment restoration
- MASTER status restoration to `παρών`

## K build failure

Focused source gate passed 62 tests, but K failed before artifact creation during ledger installation:

`DAILY_INPUT ledger refresh marker not found`

Root cause: canonical builder order is:
1. workload patch,
2. operator visibility/status normalization patch,
3. daily-change ledger patch.

The ledger installer still searched for the old refresh tail:

`ApplyMasterView -> RefreshOperationalViews = True`

while the status patch had already changed it to:

`ApplyMasterView -> NormalizeMasterPatientStatusPresentation master -> RefreshOperationalViews = True`

Fix:
- ledger patcher now supports both valid forms,
- in canonical chained form the ledger refresh is inserted after status normalization.

Relevant commits:
- `6eca6b5` ledger refresh compatibility
- `f56a9c5` canonical chained-patch regression
- `eaff5e2` K failure checkpoint

K filename is retired. No K workbook exists.

## L build

`UNIFIED_CLOSE_DAY_RECOVERY_20261031_L.xlsm` built successfully.

Evidence:
- production source unchanged: `True`
- application contract: `24`
- SHA-256 starts `E10193F84BC080E968C03935A89FF8694321F781C82...`
- size: `470409` bytes
- last write: `2026-10-31 18:28:52` local

Build checkpoint commit:
- `cc985c8`

## L native restore finding

The restore transaction itself persisted the requested changes, but the central operational view then reported:

`Οι αλλαγές αποθηκεύτηκαν, αλλά η κεντρική εικόνα δεν ενημερώθηκε σωστά.`

Diagnostic log `%TEMP%\eka_daily_input_last_error.log` showed:

- phase: `refresh operational views`
- sheet: `MASTER_SCHEDULE`
- error: `MASTER status column not found`

This proved the write/save path had succeeded and the failure was isolated to post-save refresh.

Root cause: `NormalizeMasterPatientStatusPresentation` searched only row 1 for `Κατάσταση`, but the real MASTER layout has its header row elsewhere.

### L refresh fix

The helper was changed to follow the same layout-aware strategy as the Python MASTER overlay:

1. scan the upper MASTER area for the header row identified by `Θάλαμος` in column B and `Ασθενής` in column C,
2. search that row for `Κατάσταση`,
3. normalize non-empty status cells to `Font.Strikethrough=False`,
4. if the status column is absent, `Exit Sub` safely instead of raising a hard error.

This means missing/variant status layout can no longer abort the whole operational refresh.

Relevant commits:
- `bb78577` layout-aware, fail-soft MASTER status normalization
- `f03c322` regression for native L `MASTER status column not found`
- `8b796c2` L refresh failure checkpoint

L is frozen as evidence. Do not continue testing on L.

## Current source gate

After the L fix, the focused source suite passed:

`60 passed, 2 warnings in 7.24s`

Warnings are the known Python 3.14/openpyxl ZipFile destructor warnings from `Tests/test_master_daily_overlay.py`:

`PytestUnraisableExceptionWarning: ValueError: I/O operation on closed file.`

They are not current gate failures, but should remain tracked for later cleanup.

## Current status

Patient/session restore stack now has:

- authoritative ledger records from DAILY_INPUT,
- per-row restore,
- selected restore backend support,
- all-supported restore backend support,
- atomic source-clear + restore marker + patch payload,
- rollback protection,
- zero-change/final-clean-day restore support,
- MASTER treatment restore support,
- MASTER status restore support,
- post-view status readability normalization,
- canonical builder compatibility with the ledger patch.

Still not native-proven end-to-end after the latest L refresh fix.

## Immediate next step

Build a fresh artifact using the next unused filename:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_M.xlsm`

Do not reuse J, K, or L.

Native M smoke order:

1. open only M,
2. enter one patient absence in DAILY_INPUT,
3. apply DAILY_INPUT,
4. verify MASTER treatment remains visible, red and struck through,
5. verify MASTER status/reason is readable and not struck through,
6. verify ledger row is visible under THERAPIST_DAILY,
7. restore only that row,
8. verify exact DAILY_INPUT A:E source row clears,
9. verify THERAPIST_DAILY returns to recurring/base display,
10. verify MASTER treatment returns to normal,
11. verify patient status returns to base state (`παρών` for the current specimen),
12. verify ledger row remains visible as `Επαναφέρθηκε`,
13. verify the same workbook remains open with no close/reopen.

Only after single-row restore passes should `Επαναφορά επιλεγμένων` and `Επαναφορά όλων` be tested.

Do not test therapist-absence restore or close-day yet.
