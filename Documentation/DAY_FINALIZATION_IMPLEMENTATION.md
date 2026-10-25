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

`Python/rehab_core/day_finalization.py` provides:

- `DayFinalizationError`;
- `FinalizedSnapshotStatus` with `CURRENT_FINAL` and `SUPERSEDED`;
- immutable `FinalizedDaySnapshot`;
- `build_finalized_day_snapshot(...)`.

The builder validates the requested service date and refuses to create a snapshot when any session remains unresolved. On success it freezes the current final counts for final/performed sessions, patient absences, cancellations, no-shows, replacements, treatment suspensions, approved exceptions and optional planned-session comparison context.

## Snapshot persistence

`Python/rehab_excel/finalized_day_snapshots.py` owns `FINALIZED_DAY_SNAPSHOTS`.

Technical rules:

- `xlSheetVeryHidden` and protected;
- authoritative header validation on every read/write;
- native Excel/COM mutation only;
- read-only openpyxl only for inspection/verification;
- only `CURRENT_FINAL` may be appended by ordinary close;
- persisted finalized snapshot requires `unresolved_sessions == 0`;
- duplicate `SnapshotId` with different content is rejected;
- exact duplicate is idempotent;
- one `CURRENT_FINAL` per service date;
- second current-final requires the future explicit amendment flow;
- COM values are read back semantically before `Save`.

The writer deliberately does not depend on `Range.NumberFormat`. Native Excel rejected that assignment in the first smoke, while `Value2` plus semantic read-back already provides the required safety.

## Therapist/day historical workload

`Python/rehab_core/therapist_day_history.py` introduces immutable `TherapistDayHistoryRow` records derived from finalized effective daily state.

The contract explicitly keeps these metrics separate:

- `planned_occupied_timeslots`;
- `planned_sessions`;
- `final_occupied_timeslots`;
- `final_active_sessions`;
- `replacements_received`;
- `released_by_cancellation_or_no_show`;
- `treatment_suspensions_affecting_load`;
- `approved_shared_or_over_capacity_events`;
- `capacity_timeslots`.

Critical semantic: a legal shared slot with two patients is **1 occupied timeslot + 2 active sessions**.

Replacement workload moves to the effective therapist/time. Student replacement work is not misreported as therapist workload. Non-performed outcomes remain attributable to the original therapist whose planned load was affected. Workload history also fails closed when unresolved sessions remain.

`Python/rehab_excel/therapist_day_history_store.py` persists these rows in `THERAPIST_DAY_HISTORY` as a separate `VeryHidden` protected technical table keyed by `(SnapshotId, TherapistId)`, with schema validation, duplicate protection, `Value2` writes and exact COM read-back before save.

## Native preview smoke wiring

`Python/tools/build_finalized_day_snapshot_preview.py` reconstructs one concrete day from authoritative operational inputs:

`recurring/outpatient base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState`

Then it:

1. builds the immutable finalized snapshot;
2. builds therapist/day historical rows from the same final state;
3. copies the input `.xlsm` to a fresh preview output;
4. appends and verifies the finalized snapshot;
5. appends and verifies therapist/day history rows;
6. rechecks the application VBA contract;
7. verifies the source workbook hash is unchanged;
8. deletes the output on persistence/verification failure.

DAILY_INPUT therapist identity is resolved against the official `SETTINGS` therapist registry as well as materialized providers. Legacy `φοιτ N` placeholders are excluded from therapist history. Current therapist capacity history uses the locked operational `6` timeslot baseline until a per-therapist capacity registry is introduced.

## Native verified checkpoint

Using `Excel/previews/DOUBLE_BOOKING_REJECTED.xlsm`, the first finalized snapshot smoke was verified successfully:

- Date: `2026-10-26`
- Snapshot: `finalized:2026-10-26:v1`
- Version/status: `1 / CURRENT_FINAL`
- Final sessions: `79`
- Replacements: `0`
- Patient absences: `0`
- No-shows: `0`
- Treatment suspensions: `1`
- Approved exceptions: `0`
- Unresolved sessions: `0`
- Input unchanged: `True`
- application contract source: `54` components
- application contract output: `55` components

This verifies the finalized snapshot header path end-to-end. The next native gate extends the same preview to therapist/day history rows.

## Regression coverage

- `Tests/test_day_finalization.py`
- `Tests/test_finalized_day_snapshots.py`
- `Tests/test_therapist_day_history.py`
- `Tests/test_therapist_day_history_store.py`
- `Tests/test_build_finalized_day_snapshot_preview.py`

The workload tests cover shared-slot semantics, replacement transfer, student exclusion, non-performed attribution and unresolved fail-closed behavior.

## Next slice after therapist/day native gate

After therapist/day persistence is native-green:

1. connect the backend to the explicit application action `Κλείσιμο ημέρας`;
2. add final close confirmation/summary UI without exposing technical sheets;
3. persist append-only finalization events;
4. begin patient/day finalized session-outcome history so patient, therapist and department reports reconcile from one truth;
5. implement explicit amendment/version-chain operation (`SUPERSEDED` -> new `CURRENT_FINAL`) without destructive overwrite.

## Safety

- baseline workbook remains immutable;
- do not save production/staging `.xlsm` with openpyxl;
- native smoke uses a fresh preview copy;
- finalization failure must produce no partial historical record in production design;
- preview output is disposable and deleted on verification failure;
- later corrections append/version history rather than editing old history in place;
- no ordinary close path may silently replace an existing `CURRENT_FINAL` snapshot.
