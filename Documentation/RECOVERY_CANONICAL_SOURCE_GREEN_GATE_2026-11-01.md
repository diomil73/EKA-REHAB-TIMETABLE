# Recovery canonical source green gate — 2026-11-01

## Status

Focused canonical source gate is GREEN.

Result reported from local Windows/Excel development environment:

- 71 passed
- 2 warnings
- runtime: 6.69s

The two warnings are the already-known Python 3.14 / ZipFile cleanup warnings and are not treated as functional failures for this gate.

## Canonical recovery changes now in the normal build path

The recovery fixes proven necessary by the frozen M/R native evidence are no longer dependent on the post-build recovery wrapper.

### MASTER rich-text strike contract

Canonical operator build integration now applies and verifies strike-through at character level for MASTER cells that may contain mixed formatting.

This specifically addresses the native R evidence:

- `MASTER_SCHEDULE!D8`
- payload intent: `presentation`
- strike token: `1`
- source tag: `master_daily_non_active_overlay`
- no text-run token
- prior native failure: `DAILY_INPUT read-back mismatch` during `verify patch`

Patient status/reason presentation is also normalized at character level so reason text such as `διακομισμένος` must remain readable and not inherit treatment strike-through.

### DAILY_CHANGE_LEDGER / restore

Canonical ledger source now includes the bounded fixes previously exercised only through the recovery wrapper:

- ledger view is anchored below the true used range instead of column A only;
- previous ledger UI is cleared only when its own `Αλλαγές ημέρας` header is present;
- restore uses the public validated `ApplyValidatedDailyInputPayload` path;
- `changeId` is explicitly declared;
- no dependency on retired `ApplyPayloadFileInPlace` remains.

## Retired path

The post-build native recovery builder/patcher used for N through R is retained only as forensic evidence. It is RETIRED and must not be used for the next candidate.

## Next gate

Build one fresh workbook using only the canonical builder:

`Python/tools/build_unified_preview.py`

The next candidate is `UNIFIED_CLOSE_DAY_RECOVERY_20261101_S.xlsm`.

Required native smoke remains narrowly scoped to one patient absence and one row restore. Do not test therapist-absence restore, restore-all, or `CloseOperationalDay` yet.
