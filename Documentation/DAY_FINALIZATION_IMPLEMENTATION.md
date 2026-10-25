# Day Finalization Implementation

Last updated: 2026-10-25

## Purpose

This document tracks the implementation of the explicit `Κλείσιμο ημέρας` workflow and the durable finalized-history foundation defined in `Documentation/STATISTICS_HISTORY_DESIGN.md`.

## Locked contract

A service day becomes historical truth only through an explicit close action. Finalization is fail-closed:

- every supplied `DailySessionState` must belong to the selected service date;
- no session may remain in `THERAPIST_ABSENT` / replacement-needed state;
- unresolved sessions must be zero;
- counts are derived from the final effective daily state, never reconstructed from a later recurring schedule;
- the finalized snapshot object is immutable;
- audit timestamps are timezone-aware;
- physical workbook persistence lives in hidden/protected technical tables inside the same `.xlsm`.

## Implemented core slice

`Python/rehab_core/day_finalization.py` provides `DayFinalizationError`, `FinalizedSnapshotStatus`, immutable `FinalizedDaySnapshot`, and `build_finalized_day_snapshot(...)`.

The builder refuses to create a snapshot when any session remains unresolved. On success it freezes final/performed sessions, patient absences, cancellations, no-shows, replacements, treatment suspensions, approved exceptions and optional planned-session context.

## Snapshot persistence

`Python/rehab_excel/finalized_day_snapshots.py` owns `FINALIZED_DAY_SNAPSHOTS`.

Rules:

- `xlSheetVeryHidden` and protected;
- native Excel/COM mutation only;
- semantic COM read-back before save;
- finalized snapshots require `unresolved_sessions == 0`;
- exact duplicates are idempotent;
- conflicting duplicate IDs are rejected;
- one `CURRENT_FINAL` per service date;
- later correction must use the explicit amendment/version chain.

The writer deliberately does not depend on `Range.NumberFormat` because native Excel rejected that unnecessary assignment during the first smoke.

## Therapist/day historical workload

`Python/rehab_core/therapist_day_history.py` provides immutable `TherapistDayHistoryRow` records from finalized effective daily state.

Persisted metrics include planned occupied timeslots/sessions, final occupied timeslots, final active sessions, replacements received, load released by cancellation/no-show, treatment suspensions, approved shared/over-capacity events and capacity context.

Critical semantic: a legal shared slot with two patients is **1 occupied timeslot + 2 active sessions**.

`Python/rehab_excel/therapist_day_history_store.py` persists rows in `THERAPIST_DAY_HISTORY`, keyed by `(SnapshotId, TherapistId)`, with `VeryHidden` protection, schema validation and COM read-back.

## Append-only day-history events

`Python/rehab_core/day_history_events.py` defines immutable event records and `DAY_FINALIZED`.

`Python/rehab_excel/day_history_event_store.py` persists append-only events in `DAY_HISTORY_EVENTS` with event ID, service date, event type, timezone-aware timestamp, actor/source, optional therapist/patient/session/reason links, related event/snapshot links and concise summary.

Duplicate event IDs are rejected. Ordinary close never rewrites old events.

## Explicit close-day application action

`Python/tools/close_day.py` is the backend contract for `Κλείσιμο ημέρας`.

Production ordering is:

1. build and verify finalized snapshot;
2. build and verify therapist/day history from the same state;
3. append exactly one `DAY_FINALIZED` event;
4. reread the event exactly;
5. recheck VBA application contract;
6. verify source hash unchanged;
7. expose verified preview + source hash to authoritative commit;
8. delete preview on failure.

`Python/rehab_excel/close_day_action_vba.py` provides `CloseOperationalDay` with explicit Yes/No confirmation and authoritative commit/reopen flow.

The application shell now installs and verifies `modCloseDayAction`. The MASTER toolbar now exposes **`Κλείσιμο ημέρας`**, routed through `ToolbarCloseDay -> CloseOperationalDay`. The pre-clean path also removes any prior `modCloseDayAction`, so reinstall remains idempotent.

## Native verified checkpoints

Using `Excel/previews/DOUBLE_BOOKING_REJECTED.xlsm`:

### Snapshot V1

- Date: `2026-10-26`
- Snapshot: `finalized:2026-10-26:v1`
- Final sessions: `79`
- Treatment suspensions: `1`
- Unresolved sessions: `0`
- Input unchanged: `True`
- application contract: `54 -> 55`

### Snapshot + therapist/day V2

- Final sessions: `79`
- Therapist/day history rows: `16`
- Unresolved sessions: `0`
- Input unchanged: `True`
- application contract: `54 -> 56`

### Close-day backend V1

- Snapshot: `finalized:2026-10-26:v1`
- Event: `day-finalized:2026-10-26:v1`
- Final sessions: `79`
- Therapist/day history rows: `16`
- Unresolved sessions: `0`
- Input unchanged: `True`
- application contract: `54 -> 57`

This closes the backend snapshot + therapist/day + finalization-event path end-to-end.

## Regression coverage

- `Tests/test_day_finalization.py`
- `Tests/test_finalized_day_snapshots.py`
- `Tests/test_therapist_day_history.py`
- `Tests/test_therapist_day_history_store.py`
- `Tests/test_build_finalized_day_snapshot_preview.py`
- `Tests/test_day_history_events.py`
- `Tests/test_day_history_event_store.py`
- `Tests/test_close_day_cli.py`
- `Tests/test_close_day_action_vba.py`
- `Tests/test_master_toolbar_vba.py`

## Exact next gate

1. run focused MASTER/close-action/installer tests;
2. build a fresh application-shell preview from `CLOSE_DAY_V1.xlsm` using `build_registration_menu_preview.py`;
3. require `modCloseDayAction` and the `Κλείσιμο ημέρας` MASTER button;
4. native-click the disposable preview button and verify confirmation, backend close, authoritative commit, automatic reopen to MASTER and no source corruption;
5. then begin patient/day finalized session-outcome history;
6. later implement explicit amendment/version-chain operation (`SUPERSEDED` -> new `CURRENT_FINAL`).

## Safety

- baseline immutable;
- no production/staging `.xlsm` save with openpyxl;
- native smoke uses fresh preview copies;
- failure must not create partial finalized history;
- authoritative replacement occurs only after verified preview + source-hash match;
- corrections append/version history rather than rewriting old finals;
- ordinary close may never silently replace an existing `CURRENT_FINAL` snapshot.
