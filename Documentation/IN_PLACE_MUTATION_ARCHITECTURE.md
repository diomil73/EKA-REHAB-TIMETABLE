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

## What may still use copy workflows

Copies remain valid for controlled non-production work and for rollback/recovery only:

- development previews;
- native smoke tests;
- installers/upgraders;
- migration tooling;
- backup/recovery tooling;
- diagnostics for administrators/developers.

A temporary `SaveCopyAs` recovery file is acceptable inside a production action only as rollback protection while the live workbook remains open and authoritative. It must never replace the live workbook during the normal success path.

## Proven production patterns

### Confirmed double-booking

`Python validate-only -> verified response payload -> VBA writes into already-open workbook -> read-back -> Save -> continue`

### `Κλείσιμο ημέρας`

`Python/tools/prepare_close_day_in_place.py` computes a verified snapshot + therapist-history + event payload without mutating the workbook. `modCloseDayAction` preflights all three technical stores, writes all records in place, rereads them, rolls back on failure, saves once, and remains on MASTER.

### `DAILY_INPUT`

`Python/tools/prepare_daily_input_in_place.py` builds the validated daily `WritePlan` without mutating the workbook. `modDailyInputAction` applies the serialized patches in place, verifies values/presentation, uses a temporary `SaveCopyAs` recovery copy for rollback only, saves the same open workbook, and routes to replacements only when needed.

## Atomicity expectation

Where one user action writes multiple technical records or multiple cells, production must avoid leaving a partially committed logical action. Use staged validation first, deterministic write ordering, read-back, and compensating rollback.

For day finalization specifically, snapshot + therapist/day history + `DAY_FINALIZED` event are one logical transaction and must either all be committed consistently or the action must fail without presenting the day as closed.

For DAILY_INPUT, the original workbook state is recoverable from the temporary rollback copy if any patch/read-back fails before the final successful save.

## Regression guard

`Tests/test_production_actions_stay_in_session.py` rejects close/reopen/reopen-worker mechanics in the currently migrated production action modules.

Focused module tests additionally forbid the retired mechanics and assert hidden Python execution plus operator-facing messages without implementation details.

As more production action modules are added or migrated, add them to the shared guard.
