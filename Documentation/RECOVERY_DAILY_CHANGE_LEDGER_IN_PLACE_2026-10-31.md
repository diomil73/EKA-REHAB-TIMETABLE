# Recovery checkpoint: DAILY change ledger + in-place restore

Date: 2026-10-31
Branch: `work/history-statistics-rebuild`

## Evidence entering this slice

Focused source gate passed: **43 passed, 2 warnings in 5.02s**.

This green gate covered the read-only daily-change restore preparer, operator visibility, MASTER daily overlay, DAILY_INPUT in-place preparation/action, and production in-session safety.

Native H remains frozen as evidence. No `CloseOperationalDay` is permitted on H.

## User-facing contract recovered from historical material

The established daily-change UI belongs under `THERAPIST_DAILY` and must preserve history rather than erase it.

Required presentation:

- one row per daily operational change;
- columns for subject, `Από`, `Προς`, `Αιτία`, `Κατάσταση`;
- active rows have a checkbox and per-row `Επαναφορά` action;
- bulk actions `Επαναφορά επιλεγμένων` and `Επαναφορά όλων`;
- restored rows remain visible as history with status `Επαναφέρθηκε`.

The old close/reopen or preview-sync restore backend is not being restored as the primary architecture.

## New bounded implementation

### MASTER reason readability

`430c2e90f65ee97cfd28577dd47af0aec56fd58f`

The daily patient status/reason overlay is explicitly non-struck-through and its readback verifies that property.

### Read-only restore preparation

`8573572df78e1d3cff6d1847f04a3c3e9ffdedca`

Added `Python/tools/prepare_daily_change_restore_in_place.py`.

It is read-only and currently supports:

- `patient_absence`;
- `session_cancellation`.

It recomputes authoritative daily truth with only the selected source changes ignored, preserves all other DAILY_INPUT facts and active replacements, rebuilds THERAPIST_DAILY + MASTER projections, validates the write plan, and emits a restore payload.

Therapist-absence restore is intentionally fail-closed until related replacement facts can be rolled back coherently.

### Atomic multi-change preparation

`9ceed14b8cf952526f1d93764913e312a609101d`

The restore preparer accepts repeated `--change-id` arguments and computes one final state for all selected changes. This is the backend contract for `Επαναφορά επιλεγμένων` and `Επαναφορά όλων`.

### DAILY_INPUT payload owns ledger records

`5a239a8886b2d3138007f3b5fc775ea144d2c0b5`

`prepare_daily_input_in_place.py` now serializes authoritative `CHANGE` rows before projection `PATCH` rows. Each change carries:

- immutable change id;
- date;
- type;
- subject id/display name;
- human-readable from/to text;
- reason;
- source row.

This avoids a second source of truth for the ledger.

### Native in-session ledger and restore action

`63247d647ea39b2b412a1bd3c52b399f104ebafc`

Added `Python/rehab_excel/daily_change_ledger_in_place_vba.py` with canonical module `modDailyChangeLedger` and hidden sheet `DAILY_CHANGE_LEDGER`.

The module provides:

- ledger rendering under `THERAPIST_DAILY`;
- row checkboxes;
- per-row restore;
- selected restore;
- all restore;
- hidden Python worker invocation;
- in-session application of validated restore patches;
- rollback-copy recovery on native failure;
- `%TEMP%\eka_daily_restore_last_error.log` diagnostics;
- restored rows retained as historical rows.

Normal success path does not close/reopen the active workbook and does not synchronize a generated preview back into it.

### Canonical build / application contract

`6fe58bd8cced15a8e5732f40ba81168622aefee6`

The registration/operational build installs `modDailyChangeLedger` after the operator visibility patch and creates/normalizes the hidden ledger sheet.

`ba5961c84b741db78ad797c579661c1dea510a71`

The application contract now requires **24 VBA components** and requires `DAILY_CHANGE_LEDGER` with the exact 12-column schema.

### Regression coverage

- `822f28271bd45038194a11237a56fdd83c99b878`: atomic multi-change preparer coverage.
- `6644b36d3116cea19a5c84876b02748de27ba4ec`: 24-component + ledger schema application contract.
- `3b1642a04a9be9cccfdf4b9671a003dbb44f2b67`: ledger UI, hidden in-session restore, rollback, builder wiring.
- `7fa176142f545af01cb3b08f5cdd25b6175a5ae3`: authoritative `CHANGE` rows emitted with DAILY_INPUT payload.

## Explicit current limitation

Therapist-absence and replacement-fact restore are visible in the historical ledger direction but are **not yet enabled for restore** by the new worker. The UI must not imply that they are safely restorable until coordinated rollback of related ACTIVE replacement facts is implemented and tested.

This limitation does not block proving patient-absence restore first.

## Next gate

Before any fresh native artifact I is built, run the focused source gate covering:

- multi-change restore preparation;
- DAILY_INPUT change-record serialization;
- native ledger/restore wiring;
- operator patch compatibility;
- application contract 24;
- canonical builder;
- production in-session safety.

If green, build a fresh `UNIFIED_CLOSE_DAY_RECOVERY_20261031_I.xlsm`. Do not overwrite H or any previous recovery artifact.
