# Recovery In-Session Gate — 2026-10-30

## Operator-confirmed green gate

`63 passed in 3.92s`

This source/build gate confirms the recovered normal production path after native smoke B exposed the old close/reopen DAILY_INPUT handoff.

Covered areas:

- in-session `modDailyInputAction`,
- in-session `modGlobalReplacementAction`,
- production actions remain in-session,
- read-only DAILY_INPUT preparation,
- read-only global replacement preparation,
- pure daily suspension resolution,
- special-service replacement mapping,
- reclining adjacency preference,
- strengthened application contract including `modGlobalReplacementAction`,
- outpatient registration-menu diagnostic patch,
- unified builder regressions,
- close-day VBA action regressions.

The gate explicitly rejects normal-production reintroduction of:

- `ThisWorkbook.Close`,
- `Application.OnTime`,
- `authoritative_commit_worker_cli.py`,
- `post_reopen_target`,
- console-backed `.Exec(` paths.

## Safety boundary

- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` remains frozen native-smoke evidence and must not be overwritten or repaired in place.
- Production workbook remains immutable.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm` is frozen as the current native-smoke evidence after the results below and must not be repaired in place.
- No close-day native test is allowed until outpatient shell and DAILY_INPUT smoke are healthy on a later fresh copy.

## Fresh C unified build

Operator-confirmed canonical build of `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm` succeeded.

Build output confirmed every unified stage:

- registration menu + forms,
- patient registration bridge,
- patient edit flow,
- outpatient schedule form,
- outpatient conflict guard,
- DAILY_INPUT sheet,
- MASTER projection refresh for 93 patients,
- final user navigation,
- MASTER visual baseline.

Final builder verification reported:

- `UNIFIED PREVIEW OK`,
- `DAILY_INPUT date: 2026-10-30`,
- `Source unchanged: True`,
- VBA project present,
- DAILY_INPUT sheet present,
- application contract verified.

Independent post-build verification reported:

- `SOURCE HASH UNCHANGED: True`,
- application contract count: `21`,
- `modGlobalReplacementAction: True`,
- `modCloseDayAction: True`,
- `modTherapistWorkloadStrip: True`.

Artifact metadata reported by the operator:

- path: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm`,
- size: `424713` bytes,
- last write: `30/10/2026 11:55:40 μμ`,
- SHA256 begins `56C66E90E592E07DDA9F5D67AFEF6A132842B790857...`.

## Native C smoke findings

Operator smoke on C exposed two remaining native failures before close-day testing:

1. Selecting outpatient schedule management displayed `Η διαχείριση προγραμμάτων εξωτερικών ασθενών δεν είναι διαθέσιμη σε αυτό το αρχείο.`
2. DAILY_INPUT with one patient absence displayed `Η ημερήσια κατάσταση δεν εφαρμόστηκε. Ελέγξτε τις καταχωρίσεις και δοκιμάστε ξανά...`.

Important progress relative to B:

- the C workbook remained open during the DAILY_INPUT failure,
- the previous `Close of object '_Workbook' failed` symptom did not recur,
- therefore the recovered in-session architecture removed the old close/reopen failure path even though preparation/apply still failed for a separate reason.

Outpatient source inspection established that the reported message comes from `REGISTRATION_MENU_OUTPATIENT_CLICK` in `outpatient_schedule_manager_vba.py`, which still converts any runtime failure from `ShowOutpatientScheduleManager` into the same generic unavailable message. The manager initializes by reading authoritative sheet `OUTPATIENT_SCHEDULE`; this sheet name matches `outpatient_schedule_source.py`, so there is no evidence of a stale sheet-name contract. The next diagnosis must identify the actual manager initialization/runtime error rather than rewrite scheduling logic.

DAILY_INPUT now writes a hidden diagnostic file `%TEMP%\eka_daily_input_last_error.log` containing phase, patch index, sheet, cell, error number and error text. The next diagnosis must use that log before any code change.

## Next exact step

1. Do not run `CloseOperationalDay` on C.
2. Read `%TEMP%\eka_daily_input_last_error.log` after the failed C DAILY_INPUT attempt.
3. Read-only inspect C for existence and headers of `OUTPATIENT_SCHEDULE`.
4. Use those two facts to make the next bounded source fix.
5. Any later native experiment must use a new uniquely named workbook, not B or C.
