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
- native in-place close-day UI: VERIFIED. First click finalized successfully while workbook stayed open on MASTER; no close/reopen path was required.
- native second-click duplicate protection: VERIFIED. No duplicate history was written. Initial generic failure wording was corrected in code to the operator message `Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.`
- focused combined in-place regression after migration: `57 passed in 2.65s`.

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

Temporary copies are allowed only for development/smoke/install/migration or rollback/recovery while the live workbook remains open and authoritative.

### Diagnostic-first native rule

Every production/user-facing action that can fail during native Excel execution must have a hidden developer diagnostic trace **before the first native smoke test**.

Do not repeat blind native smokes after a generic operator failure. Read the diagnostic trace first, fix the exact phase/cell/error, then rebuild a fresh disposable workbook.

Normal users continue to see only operational guidance. Diagnostic files live under `%TEMP%` and are developer/admin troubleshooting aids only.

## `Κλείσιμο ημέρας` production state

Historical model/backend and in-place UI are native-verified.

Implemented and verified:

- `Python/tools/prepare_close_day_in_place.py` performs read-only validation/calculation;
- emits one payload containing `SNAPSHOT`, all `THERAPIST` rows and one `EVENT` row;
- rejects already-finalized dates and duplicate historical keys before any workbook write;
- `modCloseDayAction` contains no workbook close/reopen/reopen-worker logic;
- VBA preflights all three schemas/duplicates;
- VBA writes snapshot + therapist history + event into the already-open workbook;
- every inserted row is reread before Save;
- failure rolls back inserted rows/new sheets;
- success protects/hides technical sheets, saves once and remains on `MASTER_SCHEDULE`;
- second click is idempotent and reports `Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.`;
- VBA source now normalizes Python raw-string backslashes before module installation so runtime paths/escape parsing use one separator.

## `DAILY_INPUT` production state

The old preview -> authoritative commit -> close/reopen production path is retired.

Implemented:

- `Python/tools/prepare_daily_input_in_place.py` rebuilds the dated effective state and validated `WritePlan` without mutating the workbook;
- official therapist registry names are supplied to `read_daily_input`;
- existing ACTIVE `DAILY_REPLACEMENT_FACTS` are included in `build_daily_session_states`;
- `replacement_sessions_needed` is calculated after active replacements are applied;
- payload contains only validated cell patch semantics;
- `modDailyInputAction` contains no `ThisWorkbook.Close`, `Application.OnTime`, authoritative commit worker or reopen target;
- VBA applies patches directly to the live workbook;
- each patch is read back before final save;
- a temporary `SaveCopyAs` file is rollback-only;
- success keeps the workbook open and routes to replacements only when replacement work remains, otherwise MASTER.

### Native DAILY_INPUT smoke status

The first true diagnostic-enabled failure was captured exactly:

```text
phase=create rollback
patch_index=12
sheet=THERAPIST_DAILY
cell=E8
error_number=1004
error=... TEMP\\eka_daily_input_rollback_...xlsm ... not found
```

Root cause: Python raw-string source had doubled backslashes inside VBA string literals. VBA does not use backslash escaping, so runtime paths contained `\\` and `SaveCopyAs` failed before any patch write.

Fixes now on branch:

- `DAILY_INPUT_MODULE_CODE` normalizes doubled raw-string backslashes to single VBA runtime backslashes before installation;
- the same normalization also fixes the one-character payload escape decoder;
- mixed/Null `Font.Size` remains handled safely;
- rich-text runs remain verified individually;
- `%TEMP%\eka_daily_input_last_error.log` captures phase, patch index, sheet, cell and VBA error on failure;
- `CLOSE_DAY_MODULE_CODE` receives the same separator normalization proactively;
- shared regression now rejects doubled runtime VBA backslashes in migrated production actions.

## Shared production regression

`Tests/test_production_actions_stay_in_session.py` guards close-day and DAILY_INPUT against reintroduction of:

- `ThisWorkbook.Close`;
- `Application.OnTime` handoff;
- authoritative commit worker;
- reopen targets/reopen flags;
- doubled runtime VBA backslashes caused by Python raw-string escaping.

`Tests/test_prepare_daily_input_in_place.py` guards that active replacement facts are loaded into the effective DAILY_INPUT state before presentation/routing is computed.

`Tests/test_daily_input_action_vba.py` guards mixed-font handling and hidden diagnostics.

Both actions must run Python hidden and keep operator messages free of implementation details.

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

A legal shared slot with two patients is **1 occupied timeslot + 2 active sessions**. Student replacement work is not silently counted as therapist workload.

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

**Do not return to close/reopen mechanics. Close-day in-place is native green.**

Immediate order:

1. pull latest backslash normalization + regression guard;
2. run focused DAILY_INPUT/close-day production-action tests;
3. build fresh `DAILY_INPUT_IN_PLACE_V4.xlsm` from `DOUBLE_BOOKING_REJECTED.xlsm`;
4. native test DAILY_INPUT once;
5. if it fails, read `%TEMP%\eka_daily_input_last_error.log` immediately before any second smoke;
6. once DAILY_INPUT is native green, continue patient/day finalized outcome history;
7. later implement explicit finalized-day amendment/version chain.

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
