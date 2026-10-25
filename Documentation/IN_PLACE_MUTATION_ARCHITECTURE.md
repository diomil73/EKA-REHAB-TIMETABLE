# In-place production mutation architecture

Last updated: 2026-10-25

## Locked product/architecture rule

No normal production/user-facing workflow may close and reopen the workbook in order to complete an action.

The workbook is the live application host. Ordinary actions must complete while the application stays open.

## Required production pattern

For any user-facing mutation:

1. validate the requested business action;
2. compute the authoritative result;
3. write only the required cells/tables into the already-open workbook using native Excel/VBA/COM-compatible in-session mutation;
4. immediately read back the written values semantically;
5. if read-back fails, roll back the just-written row/range where feasible and report failure;
6. save the same open workbook;
7. remain in the application shell on the appropriate operational screen.

No production action may depend on replacing the open `.xlsm` file from a preview copy.

## What is forbidden in production UI paths

- `ThisWorkbook.Close` as part of completing an ordinary action;
- automatic reopen of the workbook after an action;
- file-lock waiting as part of normal operator flow;
- verified-preview -> replace authoritative workbook -> reopen transactions;
- requiring the user to close/reopen Excel to complete an action;
- console windows or technical process messages.

## What may still use preview/copy workflows

Preview/copy/replace mechanics remain valid only for controlled non-production purposes such as:

- development previews;
- native smoke tests;
- installers/upgraders that intentionally rebuild a disposable or staging workbook;
- migration tooling;
- backup/recovery tooling;
- diagnostics where the operator is explicitly acting as administrator/developer.

These paths must never leak into routine application actions.

## Proven reference pattern

The confirmed double-booking workflow is the current reference:

`Python validate-only -> verified response payload -> VBA writes into already-open workbook -> read-back -> Save -> continue`

This is the preferred model for future production actions.

## Known production paths requiring migration

### `Κλείσιμο ημέρας`

Current preview/authoritative-commit/deferred-close implementation is now considered transitional and must be replaced before the close-day UI is production-complete.

Target:

`validate final day -> produce snapshot/history/event payload -> in-place write all three technical stores -> semantic read-back -> Save -> remain on MASTER`

### `DAILY_INPUT`

`Python/rehab_excel/daily_input_action_vba.py` still uses preview -> authoritative commit -> close/reopen. This is now explicit migration debt.

Target:

`validate DAILY_INPUT -> compute resulting daily facts/projection -> in-place write verified changes -> read-back -> Save -> remain open / route to next operational screen`

## Atomicity expectation

Where one user action writes multiple technical records, the production implementation must avoid leaving a partially committed logical action. Use staged validation first, then deterministic write ordering, read-back, and compensating rollback of newly appended rows when any subsequent write fails.

For day finalization specifically, snapshot + therapist/day history + `DAY_FINALIZED` event are one logical transaction and must either all be committed consistently or the action must fail without presenting the day as closed.

## Regression requirement

Once the known production paths are migrated, add a repository regression guard that rejects close/reopen/file-replacement mechanics inside user-facing production VBA action modules. Developer/smoke/install/recovery tooling is excluded from that guard.
