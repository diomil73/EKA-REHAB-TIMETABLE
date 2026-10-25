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
- focused DAILY_INPUT/global replacement integration gate before V5: `77 passed in 5.33s`

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

## DAILY_INPUT native history

### V4

- patient absence applied in-place successfully;
- therapist absence incorrectly routed to raw/blank `REPLACEMENTS` sheet;
- this raw-sheet route is retired permanently.

### V5 finding: MASTER + workload were not operationally correct

Native V5 used a real patient cancellation (`ΓΙΑΝΝΙΑ`, reason `Διακομιστήριο`). The DAILY_INPUT action completed, but:

- `MASTER_SCHEDULE` still presented the patient/session as active;
- therapist `ΠΑΠΑΜΑΝΟΥ` still showed `6/6`, although the cancelled session should release one occupied timeslot.

Root causes found:

1. `master_daily_overlay.py` only emitted MASTER patches for `REPLACED` states. `PATIENT_ABSENT`, `CANCELLED` and `THERAPIST_ABSENT` were ignored entirely.
2. `modTherapistWorkloadStrip` counted every non-empty THERAPIST_DAILY cell as occupied. Cancelled/absent sessions remain visibly struck through, so the cell is intentionally non-empty and was therefore counted incorrectly.
3. The preview pre-clean knew about global replacement / double-booking / workload components, but `create_registration_menu_preview()` did not reinstall those operational components. A fixture could therefore mask missing installer coverage.

Fixes now on branch:

- MASTER overlay handles `PATIENT_ABSENT`, `CANCELLED`, `THERAPIST_ABSENT`, and `REPLACED`;
- affected treatment cells explicitly show the dated operational outcome (`ΑΠΩΝ`, `ΑΚΥΡΩΘΗΚΕ`, or `ΑΝΑΜΟΝΗ ΑΝΑΠΛΗΡΩΣΗΣ`) instead of leaving the recurring treatment visually active;
- when all dated sessions for an inpatient are non-active, the MASTER patient status column is also overlaid;
- workload badges now count a slot only when the cell contains at least one operationally active line;
- struck-through cancellation/absence lines do not count;
- italic replacement-source lines and `→` routing annotation lines do not count;
- a shared slot with one cancelled and one active patient remains exactly one occupied timeslot;
- the preview builder now explicitly installs and verifies `modGlobalReplacementAction`, `modDoubleBookingReview`, `frmDoubleBookingReview`, and `modTherapistWorkloadStrip` after the shell/menu build.

Relevant commits:

- `5288de5987edba9dd5b944c1b47005d0374d4519` MASTER non-active daily overlay
- `d727825c3af67f71ae385590797b18a12b2eb513` operational workload counting
- `40bde2fb328fc624313d38bb254bcbb091c79cfd` operational module installation in preview build
- `fdcffdac4fb7c931eab2ebb7b450fda2ee3f82eb` MASTER cancellation/absence regression tests
- `0f0aade7b52709904c89df1e0d5bf460edf8f56d` workload inactive-line regression tests
- `17fd74052b942fa377ea41a88c79c2fdb603d97c` preview operational-module build contract

## DAILY_INPUT integration architecture

### Operational projection gate

`Python/tools/prepare_daily_input_in_place.py` composes both:

- patient/therapist daily presentation;
- `MASTER_SCHEDULE` daily overlay.

Both are serialized into the same PATCH payload and therefore pass the same in-place write/read-back gate.

### Workload gate

`modDailyInputAction` calls `RenderTherapistWorkloadStrip` after successful patch application and verifies the MASTER refresh shape before reporting success.

Workload remains occupied-timeslot based: `Όνομα n/6`, not patient-line count.

### Guided/global replacement continuation

`Python/tools/prepare_global_replacements_in_place.py` reuses the authoritative global solver and includes current DAILY_INPUT facts, ACTIVE replacement facts, provider policy, service hours, patient cross-specialty availability, students, reclining adjacency and explicit double-booking review detection.

`Python/rehab_excel/global_replacement_action_vba.py` persists ordinary replacement facts, generated suspensions and final operational projection patches in-place with rollback and `%TEMP%\eka_replacement_last_error.log` diagnostics.

If an explicit double-booking decision is required, it invokes the guided `ShowDoubleBookingReview` form. It never activates a raw worksheet.

## Shared production regression

`Tests/test_production_actions_stay_in_session.py` guards close-day, DAILY_INPUT and global replacements against close/reopen workers, visible execution, reopen targets and doubled runtime VBA backslashes.

Focused integration tests now also include:

- `Tests/test_master_daily_overlay.py`
- `Tests/test_therapist_workload_strip_vba.py`
- `Tests/test_build_registration_menu_operational_modules.py`

## Exact resume point

Do not return to close/reopen mechanics and do not route to raw `REPLACEMENTS`.

Immediate order:

1. run focused MASTER/workload/DAILY_INPUT/global replacement/build tests;
2. if green, build fresh `DAILY_INPUT_IN_PLACE_V6.xlsm` from `DOUBLE_BOOKING_REJECTED.xlsm`;
3. native cancel one patient session and verify explicitly:
   - MASTER affected treatment/status no longer appears active;
   - THERAPIST_DAILY shows the cancellation/absence presentation;
   - affected therapist occupied badge drops when the slot becomes free;
4. then test therapist absence and verify global/guided replacement continuation with no raw worksheet;
5. on any native failure, read `%TEMP%\eka_daily_input_last_error.log` or `%TEMP%\eka_replacement_last_error.log` before retrying;
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
