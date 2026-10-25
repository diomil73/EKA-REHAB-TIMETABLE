# Current Handoff

Last updated: 2026-10-25

## Read order

Start with `Documentation/ONBOARDING.md`, then read `Documentation/IN_PLACE_MUTATION_ARCHITECTURE.md`. This file is the exact resume point.

## Branch / trusted state

- Branch: `feature/outpatient-scheduling`
- Baseline `Rehab_Center_System_v27_1.xlsm` remains immutable.
- Replacement-quality milestone is CLOSED end-to-end.
- Statistics/history milestone is active.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER gate: `59 passed in 6.85s`
- native student replacement smoke: VERIFIED
- native explicit REJECT branch: VERIFIED, 1 persisted `Αναστολή θεραπείας`, unresolved 0
- native explicit ACCEPT branch: VERIFIED
- native finalized-day snapshot smoke: VERIFIED on `2026-10-26`, 79 final sessions, 1 treatment suspension, unresolved 0
- native snapshot + therapist/day history smoke: VERIFIED, 16 therapist/day rows
- native close-day backend smoke: VERIFIED, snapshot + 16 therapist/day rows + one `DAY_FINALIZED` event, application contract `54 -> 57`

## Locked production interaction architecture

**No normal production/user-facing workflow may close and reopen the workbook to complete an action.**

Required model:

`validate -> compute authoritative result -> in-place native write into already-open workbook -> semantic read-back -> Save -> continue in application`

Forbidden in normal production actions:

- `ThisWorkbook.Close` as part of action completion;
- automatic workbook reopen;
- preview-copy replacement of the authoritative open workbook;
- file-lock waiting in operator flow;
- requiring the user to close/reopen Excel;
- console/Command Prompt windows;
- user-visible backend/worker/preview/hash/file-lock details.

Preview/copy/replace workflows remain allowed only for developer smoke tests, installation, migration, diagnostics and recovery tooling.

Reference implementation pattern: confirmed double-booking validate-only path.

## Close-day state

Historical model/backend remains native-verified.

The old close/reopen UI path is retired and no longer the production candidate.

Implemented now, awaiting focused tests/native smoke:

- `Python/tools/prepare_close_day_in_place.py` performs read-only validation/calculation against the saved open workbook;
- it emits one verified payload containing `SNAPSHOT`, all `THERAPIST` rows and one `EVENT` row;
- it rejects already-finalized dates and duplicate historical keys before any workbook write;
- `modCloseDayAction` no longer contains `ThisWorkbook.Close`, `Application.OnTime`, reopen target or authoritative commit worker logic;
- VBA preflights the three history schemas/duplicate keys;
- VBA writes snapshot + therapist history + event into the already-open workbook;
- every inserted row is reread before Save;
- failure triggers rollback of inserted rows/new sheets;
- success protects/hides technical sheets, saves once and remains open on `MASTER_SCHEDULE`;
- focused regression forbids reintroduction of close/reopen mechanics.

## Known production migration debt

### `DAILY_INPUT`

`Python/rehab_excel/daily_input_action_vba.py` still uses preview -> authoritative commit -> close/reopen. This is explicit migration debt and must be converted immediately after close-day native green.

Target:

`validate DAILY_INPUT -> compute dated facts/projection -> in-place verified write -> read-back -> Save -> remain open / route to next operational screen`

## Day-finalization/history state

Implemented and native-verified:

- immutable `FinalizedDaySnapshot`;
- strict unresolved=0 gate;
- `FINALIZED_DAY_SNAPSHOTS` schema;
- therapist/day historical workload derivation;
- occupied timeslots kept distinct from active sessions;
- `THERAPIST_DAY_HISTORY` schema;
- immutable append-only `DayHistoryEvent`;
- `DAY_FINALIZED` linked to snapshot;
- `DAY_HISTORY_EVENTS` schema;
- duplicate protection and semantic read-back rules;
- backend close-day reconstruction from final effective state.

The old close/reopen UI experiments and their 1004/reopen-shell fixes are regression history only.

## Therapist workload contract

MASTER badges use `Όνομα n/6`, where `n` is occupied timeslots, not patient lines:

- 0-4 green
- 5 yellow/orange
- 6+ red

A legal shared slot with two patients is one occupied timeslot and two active sessions. Student replacement work is not silently counted as therapist workload.

## Closed replacement decision chain

`recurring/base -> cancellations/no-shows -> therapist absences -> surviving ACTIVE replacements -> global solve -> double-booking review -> explicit ACCEPT/REJECT -> remaining suspension -> final recompute`

Double-booking remains explicit-only. Final operational state requires `unresolved_sessions == 0`.

## Patient-history requirements already locked

Preserve/derive:

- first operational entry date;
- end/archive date;
- stay duration;
- treatment/service days;
- scheduled/performed/missed sessions with structured reasons;
- expected working days excluding weekends/holidays;
- treatment counts and timeline;
- totals/trends by period;
- manual append-only 15-day evaluation milestones;
- aggregate patient reporting for all patients and by originating clinic.

Patient lifecycle dates are manually correctable for migration of existing patients. MASTER later shows compact start date such as `Έν. 12/09/26` near the responsible doctor where layout permits.

## Exact resume point

**Do not continue the close/reopen experiment. Do not open another reporting vertical yet.**

Immediate order:

1. run focused in-place close-day tests;
2. build a fresh application-shell preview from `DOUBLE_BOOKING_REJECTED.xlsm`;
3. native-click `Κλείσιμο ημέρας` and verify workbook stays open on MASTER, no console, no reopen, no error;
4. click `Κλείσιμο ημέρας` again and verify already-closed protection;
5. after native green, migrate `DAILY_INPUT` to in-place architecture;
6. add broader production-action regression guard against close/reopen mechanics;
7. only then continue patient/day finalized outcome history.

## Safety reminders

- baseline immutable;
- daily facts are overlays, never destructive recurring edits;
- native Excel/VBA/COM-compatible mutation for real `.xlsm` production writes;
- no production/staging `.xlsm` save with openpyxl;
- semantic read-back after production writes;
- Windows tests use `python -m pytest`;
- later corrections append/version history instead of rewriting old finals.

## Documentation rule

Every agreed design/product/future-feature idea must be recorded in repository documentation in the same work cycle.
