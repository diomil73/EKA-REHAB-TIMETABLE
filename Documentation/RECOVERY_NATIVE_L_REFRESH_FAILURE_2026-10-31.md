# Recovery native L refresh failure — 2026-10-31

## Artifact

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_L.xlsm`

Build evidence supplied by operator:
- unified preview build succeeded,
- production source hash remained unchanged,
- application contract verified with 24 required VBA components,
- artifact SHA-256 begins `E10193F84BC080E968C03935A89FF8694321F781C82...`,
- artifact size 470409 bytes,
- artifact last-write 2026-10-31 18:28:52 local.

## Native finding

During the patient-absence / restore smoke, the transactional changes were saved but the central operational view did not refresh correctly.

The authoritative `%TEMP%\eka_daily_input_last_error.log` diagnostic reported:
- phase: `refresh operational views`,
- sheet: `MASTER_SCHEDULE`,
- error: `MASTER status column not found`.

This isolates the failure to the post-`ApplyMasterView` helper `NormalizeMasterPatientStatusPresentation`, not to restore preparation, source clear, payload application, ledger persistence, or save.

## Root cause

The helper searched for the `Κατάσταση` header only on row 1. The actual MASTER layout identifies its header row by the `Θάλαμος` / `Ασθενής` pair, matching the Python MASTER overlay logic. The helper therefore raised a false hard failure during operational refresh.

## Bounded fix

`operator_visibility_patch_vba.py` now installs a layout-aware, fail-soft normalization helper:
- scan the first 12 rows for `Θάλαμος` in column B and `Ασθενής` in column C,
- search that header row for `Κατάσταση`,
- clear strike-through only on non-empty status cells below that header,
- if the header row or status column is absent, exit safely instead of raising an error.

This preserves the no-strike status invariant when supported by the workbook layout without allowing a presentation helper to block the whole operational refresh.

Regression coverage was updated in `Tests/test_operator_visibility_patch_vba.py`.

## Safety boundary

- L is frozen as native evidence and must not be repaired in place.
- production workbook remains immutable.
- no restore or close-day testing continues on L.
- the next native artifact must use a fresh filename after the focused source gate is green.
