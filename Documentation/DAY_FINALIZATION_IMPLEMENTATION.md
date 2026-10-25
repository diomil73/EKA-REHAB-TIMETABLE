# Day Finalization Implementation

Last updated: 2026-10-25

## Purpose

This document tracks the explicit `Κλείσιμο ημέρας` workflow and durable finalized-history foundation defined in `Documentation/STATISTICS_HISTORY_DESIGN.md`.

## Locked contract

A service day becomes historical truth only through an explicit close action. Finalization is fail-closed:

- every supplied `DailySessionState` belongs to the selected service date;
- no session remains replacement-needed/unresolved;
- unresolved sessions must be zero;
- counts come from final effective daily state;
- snapshot/event objects are immutable;
- timestamps are timezone-aware;
- history lives in hidden/protected technical tables in the same `.xlsm`;
- normal production close must stay inside the already-open workbook and must never close/reopen it.

## Implemented core/history slices

`Python/rehab_core/day_finalization.py` provides immutable `FinalizedDaySnapshot` and the unresolved=0 gate.

`Python/rehab_excel/finalized_day_snapshots.py` owns `FINALIZED_DAY_SNAPSHOTS`.

`Python/rehab_core/therapist_day_history.py` derives therapist/day history, preserving the distinction between occupied timeslots and active sessions. A legal shared slot with two patients is **1 occupied timeslot + 2 active sessions**.

`Python/rehab_excel/therapist_day_history_store.py` owns `THERAPIST_DAY_HISTORY`.

`Python/rehab_core/day_history_events.py` defines immutable audit events including `DAY_FINALIZED`.

`Python/rehab_excel/day_history_event_store.py` owns `DAY_HISTORY_EVENTS`.

All three technical tables are schema-checked, protected and `VeryHidden` in production use.

## Native verified backend checkpoints

Using `Excel/previews/DOUBLE_BOOKING_REJECTED.xlsm`:

- snapshot V1: 79 final sessions, 1 treatment suspension, unresolved 0;
- snapshot + therapist/day V2: 16 therapist/day rows;
- close-day backend V1: snapshot `finalized:2026-10-26:v1` + 16 therapist/day rows + event `day-finalized:2026-10-26:v1`;
- source remained unchanged in the backend preview smoke;
- application contract reached `54 -> 57` components.

These checkpoints validate the historical data model and reconstruction logic. They do not require production to use preview replacement.

## Production close-day architecture

The old `preview -> authoritative commit -> close/reopen` application flow is retired.

Production is now:

`confirm -> Save current workbook -> read-only Python validation/calculation -> verified payload -> preflight all three technical tables -> in-place write -> semantic read-back -> Save -> remain open on MASTER`

### Validation payload

`Python/tools/prepare_close_day_in_place.py`:

- reads the saved open workbook without mutating it;
- rebuilds final daily state;
- enforces unresolved=0;
- refuses an already-finalized service date;
- creates the finalized snapshot, therapist/day rows and one `DAY_FINALIZED` event in memory;
- refuses duplicate therapist-history/event keys already present;
- verifies the source hash is unchanged;
- emits a UTF-8 escaped tab-separated payload containing exactly three record kinds: `SNAPSHOT`, `THERAPIST`, `EVENT`.

### In-place application action

`Python/rehab_excel/close_day_action_vba.py` now:

- runs validation hidden;
- never calls `ThisWorkbook.Close`;
- never uses `Application.OnTime` for handoff;
- never starts the authoritative commit worker;
- preflights schemas and duplicates before writing;
- writes snapshot, all therapist rows and event into the already-open workbook;
- rereads every written row semantically before save;
- rolls back inserted rows/sheets if any write/read-back step fails;
- protects/hides the technical sheets;
- saves once after the whole logical transaction succeeds;
- stays open and activates `MASTER_SCHEDULE`.

User-visible messages contain only business confirmation, result or next-step guidance.

## Regression coverage

- `Tests/test_day_finalization.py`
- `Tests/test_finalized_day_snapshots.py`
- `Tests/test_therapist_day_history.py`
- `Tests/test_therapist_day_history_store.py`
- `Tests/test_build_finalized_day_snapshot_preview.py`
- `Tests/test_day_history_events.py`
- `Tests/test_day_history_event_store.py`
- `Tests/test_close_day_cli.py`
- `Tests/test_prepare_close_day_in_place.py`
- `Tests/test_close_day_action_vba.py`
- `Tests/test_master_toolbar_vba.py`

The close-day VBA regression explicitly forbids workbook close/reopen, `Application.OnTime`, authoritative commit worker use and reopen-target handoff.

## Exact next gate

1. run focused tests for the in-place preparation/action plus MASTER/application-shell integration;
2. build a fresh disposable application-shell workbook from `DOUBLE_BOOKING_REJECTED.xlsm`;
3. click `Κλείσιμο ημέρας`;
4. verify success while the workbook remains open on MASTER, with no console window and no reopen;
5. verify the day cannot be closed a second time;
6. after native green, migrate `DAILY_INPUT` to the same in-place architecture;
7. then continue patient/day finalized outcome history;
8. later implement amendment/version-chain correction (`SUPERSEDED` -> new `CURRENT_FINAL`).

## Safety

- baseline immutable;
- no production/staging `.xlsm` save with openpyxl;
- native smoke uses fresh preview copies;
- failure must not create partial finalized history;
- ordinary close may never silently replace an existing `CURRENT_FINAL` snapshot;
- corrections append/version history rather than rewriting old finals;
- production user actions do not close/reopen the workbook.
