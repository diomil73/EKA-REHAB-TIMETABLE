# Recovery J zero-state restore fix — 2026-10-31

## Native evidence

Frozen artifact: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_J.xlsm`.

J built successfully with the 24-component application contract and unchanged production source. Native patient-absence smoke then exposed two independent findings:

1. the MASTER `Κατάσταση` reason remained struck through after `ApplyMasterView`, despite the patch-level `Strikethrough=False` contract;
2. row restore failed during read-only Python preparation.

J remains frozen evidence. It must not be repaired, rebuilt, overwritten, or used for close-day testing.

## Exact restore diagnostic

The active supported ledger change in J was:

`daily-input:2026-10-30:patient:21`

Direct read-only execution of `prepare_daily_change_restore_in_place.py` returned:

`SAFETY STOP: At least one non-active daily state is required`

The JSON response contained the same error, and the J SHA-256 remained unchanged.

This proves the failure occurred before workbook mutation. Restoring the only active DAILY_INPUT change made the recomputed day fully ACTIVE. The general patient-centric preview correctly rejects an empty non-active state set, so it was the wrong projection path for this specific inverse operation.

## Bounded fixes

### MASTER status readability

`operator_visibility_patch_vba.py` now enforces a post-`ApplyMasterView` invariant: the dynamically located `Κατάσταση` column is never struck through. This is additional to the existing red/readable patch-level rule.

Focused source gate after this change: `27 passed, 2 known Python 3.14/openpyxl ZipFile cleanup warnings`.

### Restore to fully normal day

`prepare_daily_change_restore_in_place.py` now branches only when the recomputed post-restore truth contains no non-ACTIVE state.

Normal partial restores still use the established preview/composition path unchanged.

For a fully ACTIVE result, the new clean-day restore path:

- computes current pre-restore daily states;
- uses the current preview bindings only to identify touched `THERAPIST_DAILY` cells;
- rebuilds those exact cells from recurring base data with normal, non-struck text runs;
- derives the current MASTER overlay to identify the treatment/status cells affected by the dated change;
- rewrites treatment cells with their unchanged recurring value but normal presentation;
- restores the MASTER status value from the base master projection with normal presentation;
- builds and validates an ordinary bounded `WritePlan` using the existing write zones;
- leaves the generic preview invariant unchanged;
- remains read-only and SHA-guarded.

Commits:

- `5129b731020f027edbaacba55277fa6b9dcf9775` — clean-day restore projection path.
- `e1695cb586acff932e40f754dab04ef618d028eb` — zero-state restore regressions.

## Gate before K

Do not build K until both are true:

1. focused Python tests are green;
2. direct read-only preparation against frozen J and `daily-input:2026-10-30:patient:21` succeeds and leaves J hash unchanged.

If either fails, K is blocked and the exact source/read-only failure must be recorded here before any further change.
