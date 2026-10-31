# Recovery native K build failure — 2026-10-31

## Pre-build gate

Operator evidence:
- focused source suite passed: 62 passed, 3 warnings,
- prior read-only J restore diagnostic passed and J hash remained unchanged,
- production source hash remained unchanged during the failed K build attempt.

## K build failure

Artifact name attempted:
`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_K.xlsm`

The build failed closed during `registration menu + forms` before an output workbook was produced:

`Operational module installation failed during installing daily change ledger: DAILY_INPUT ledger refresh marker not found`

No K artifact exists. The K filename is retired and must not be reused.

## Root cause

Canonical installation order is:
1. install close-day action,
2. install therapist workload strip,
3. patch operator visibility,
4. install daily change ledger.

The operator visibility patch changes the end of `RefreshOperationalViews` from:

`ApplyMasterView -> RefreshOperationalViews = True`

to:

`ApplyMasterView -> NormalizeMasterPatientStatusPresentation master -> RefreshOperationalViews = True`.

The daily-change ledger patcher still searched only for the pre-normalization marker, so installation failed even though both transformations were individually valid.

## Bounded fix

`daily_change_ledger_in_place_vba.patch_daily_input_for_ledger` now accepts both legal refresh forms.

On the canonical path, ledger rendering is inserted after status normalization:

`ApplyMasterView -> NormalizeMasterPatientStatusPresentation master -> RefreshDailyChangeLedgerView -> RefreshOperationalViews = True`.

The generic/base path remains supported for direct unit composition.

A regression now composes the exact canonical transformation sequence and requires exactly one ledger refresh call.

## Safety boundary

- production workbook remains immutable,
- J remains frozen forensic evidence,
- K does not exist and its filename is retired,
- no runtime restore semantics changed in this fix,
- no close/reopen architecture was introduced,
- next native artifact must use a fresh filename, `L`, only after focused regression gate passes.
