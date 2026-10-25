# Current Handoff

Last updated: 2026-10-25

## Read order

Start with `Documentation/ONBOARDING.md`, then read `Documentation/IN_PLACE_MUTATION_ARCHITECTURE.md`. This file is the exact resume point.

## Branch / trusted state

- Branch: `feature/outpatient-scheduling`
- Baseline `Rehab_Center_System_v27_1.xlsm` remains immutable.
- Replacement-quality business logic is CLOSED end-to-end.
- Statistics/history milestone is active.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER gate: `59 passed in 6.85s`
- native student replacement smoke: VERIFIED
- native explicit REJECT branch: VERIFIED, 1 persisted `Αναστολή θεραπείας`, unresolved 0
- native explicit ACCEPT branch: VERIFIED
- native finalized-day snapshot smoke: VERIFIED on `2026-10-26`, 79 final sessions, 1 treatment suspension, unresolved 0
- native snapshot + therapist/day history smoke: VERIFIED, 16 therapist/day rows
- native close-day backend smoke: VERIFIED, snapshot + 16 therapist/day rows + one `DAY_FINALIZED` event
- native in-place close-day UI: VERIFIED; workbook stays open on MASTER
- native close-day second-click duplicate protection: VERIFIED
- focused combined in-place regression before V4: `57 passed in 2.65s`

## Locked production interaction architecture

No normal user-facing workflow may close/reopen the workbook to complete an action.

Required model:

`validate -> compute authoritative result -> in-place native write -> semantic read-back -> Save -> continue in application`

Forbidden in normal production actions:

- `ThisWorkbook.Close` as part of action completion;
- automatic reopen;
- preview-copy replacement of the live workbook;
- file-lock waiting;
- asking the user to close/reopen Excel;
- console/Command Prompt windows;
- backend/worker/preview/hash/file-lock language in operator messages.

Temporary copies are allowed only for development/install/migration/recovery and rollback protection while the live workbook remains open.

### Diagnostic-first native rule

Every native production action must have a hidden developer diagnostic trace before its first smoke test. After a generic operator failure, read the diagnostic before another smoke. Do not repeat blind native tests.

## `Κλείσιμο ημέρας`

Native green in-place:

- read-only Python preparation;
- snapshot + therapist history + `DAY_FINALIZED` event persisted in the open workbook;
- read-back and rollback;
- no close/reopen;
- duplicate second click reports `Η επιλεγμένη ημέρα είναι ήδη κλεισμένη.`

## `DAILY_INPUT` V4 native findings

### Patient absence

Native V4 patient-absence apply succeeded in place. The workbook stayed open and no console/reopen path appeared.

The user correctly identified that success was incomplete because the smoke had not verified:

- `MASTER_SCHEDULE` projection;
- `THERAPIST_DAILY` projection;
- therapist workload badges.

### Therapist absence

Native V4 therapist-absence apply parsed and applied, then reported that patients needed redistribution, but routed to a blank/raw `REPLACEMENTS` worksheet. This is invalid production UX.

The raw-sheet route is now retired. Normal users must never be sent to `REPLACEMENTS` for replacement work.

## DAILY_INPUT integration now implemented, awaiting focused tests + V5 native smoke

### Operational projection gate

`Python/tools/prepare_daily_input_in_place.py` now composes both:

- patient/therapist daily presentation;
- `MASTER_SCHEDULE` daily overlay.

Both are serialized into the same PATCH payload and therefore pass the same in-place write/read-back gate.

Relevant commits:

- `91fe181a126735a1c1acba4ff0860dc99dcec972` MASTER overlay in DAILY_INPUT plan
- `c36bad91792365299d473fe236307ab8150d07da` regression guard for operational projections

### Workload gate

`modDailyInputAction` now calls `RenderTherapistWorkloadStrip` after successful patch application and verifies the MASTER refresh shape before reporting success. The installer ensures `modTherapistWorkloadStrip` is present.

Workload remains occupied-timeslot based: `Όνομα n/6`, not patient-line count.

### Guided/global replacement continuation

New read-only preparation:

- `Python/tools/prepare_global_replacements_in_place.py`

It reuses the authoritative global solver and includes:

- current DAILY_INPUT facts;
- ACTIVE replacement facts;
- provider policy;
- service hours;
- patient cross-specialty availability;
- students;
- reclining adjacency;
- explicit double-booking review detection.

New in-place action:

- `Python/rehab_excel/global_replacement_action_vba.py`

It can persist:

- ordinary `DAILY_REPLACEMENT_FACTS`;
- generated treatment suspensions when no reviewable exception exists;
- final operational projection PATCH rows;
- rollback recovery and `%TEMP%\eka_replacement_last_error.log` diagnostics.

If an explicit double-booking decision is required, it invokes the existing guided `ShowDoubleBookingReview` form. It does not activate a raw worksheet.

Relevant commits:

- `09fdf9f87f5626118766188676b30dde77672010` read-only global in-place preparation
- `0abf8dd9fea0e072ba7e3ec79a3d74cd7b71dbec` preparation tests
- `2ae31fb3a31a25587f001bd1b73c15c72d0693c0` global in-place VBA action
- `06065e3aede0b824ece1602ff051aea8627ee890` DAILY_INPUT guided routing + workload refresh
- `065809161a81bb79644f0be7039b5a11981d3060` DAILY_INPUT integration tests
- `f399f937c425e6cb5b416930fc531e6a95d741de` global action tests
- `d724d704874c7bf35fe7fa09bbe3cfcb967c9476` shared production guard includes global action
- `7358de226be94a6dd34f9eb9efb1c78f817cccd9` preview pre-clean includes guided replacement/workload components
- `070ffc5b3b70c8d3328ca0bd2ccbf661a61fa192` DAILY_INPUT workflow documentation

## Shared production regression

`Tests/test_production_actions_stay_in_session.py` now guards:

- close-day;
- DAILY_INPUT;
- global replacements.

It rejects close/reopen workers, visible execution, reopen targets and doubled runtime VBA backslashes.

Additional focused tests:

- `Tests/test_prepare_daily_input_in_place.py`
- `Tests/test_daily_input_action_vba.py`
- `Tests/test_prepare_global_replacements_in_place.py`
- `Tests/test_global_replacement_action_vba.py`

## Exact resume point

Do not return to close/reopen mechanics and do not route to raw `REPLACEMENTS`.

Immediate order:

1. run focused Python/VBA source-contract tests for DAILY_INPUT + global replacement + shell/toolbar;
2. if green, build fresh `DAILY_INPUT_IN_PLACE_V5.xlsm` from `DOUBLE_BOOKING_REJECTED.xlsm`;
3. native patient-absence smoke and explicitly inspect MASTER, THERAPIST_DAILY and workload badge changes;
4. native therapist-absence smoke and verify global/guided replacement continuation, with no raw worksheet;
5. if native failure occurs, read `%TEMP%\eka_daily_input_last_error.log` or `%TEMP%\eka_replacement_last_error.log` before any retry;
6. once both branches are native green, continue patient/day finalized outcome history;
7. later implement finalized-day amendment/version chain.

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
