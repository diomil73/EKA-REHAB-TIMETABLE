# Recovery In-Session Gate — 2026-10-30

## Operator-confirmed green gates

### In-session production gate

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

### Outpatient + DAILY_INPUT recovery gate

`48 passed, 3 warnings in 3.97s`

This focused gate confirms the two bounded fixes derived from native C evidence:

- canonical creation/verification of authoritative `OUTPATIENT_SCHEDULE`,
- application-contract enforcement of the sheet and exact header schema,
- outpatient manager/build regressions,
- complete DAILY_INPUT reversible reader contract (`source_changes`, `ignore_rows`, `allow_empty`),
- in-place DAILY_INPUT preparer/action integration,
- production actions still remain in-session.

The three warnings were `PytestUnraisableExceptionWarning` messages from Python 3.14 `zipfile.ZipFile.__del__` after openpyxl/keep_vba cleanup (`ValueError: I/O operation on closed file`). They did not fail any test and did not indicate workbook mutation. Track them as cleanup noise rather than suppressing them silently, but they do not block the next native smoke candidate.

## Safety boundary

- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` remains frozen native-smoke evidence and must not be overwritten or repaired in place.
- Production workbook remains immutable.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm` is frozen as native-smoke evidence and must not be repaired in place.
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

## Confirmed outpatient structural omission

Read-only inspection of C confirmed:

- `OUTPATIENT_SCHEDULE PRESENT: False`,
- therefore no headers or rows were available,
- the manager's `UserForm_Initialize` requires `ThisWorkbook.Worksheets("OUTPATIENT_SCHEDULE")`, so the missing authoritative source sheet directly explains the generic unavailable message.

This exposed a second application-contract gap: the contract verified required VBA components but did not require the authoritative outpatient data sheet. The canonical outpatient build stage installed form/edit/manager code but did not create the underlying `OUTPATIENT_SCHEDULE` source.

Bounded recovery fixes:

- `29f74ac` add build-time `outpatient_schedule_sheet.py` initializer with exact authoritative headers and fail-closed schema verification,
- `5c82e48` invoke the initializer before outpatient edit/manager installation in the canonical build stage,
- `06ad12a` strengthen the application contract to require `OUTPATIENT_SCHEDULE` and its exact header schema,
- `b1c5196` add application-contract regressions for the required outpatient source,
- `ad4fa23` add initializer/build-order regression tests.

The initializer is build-time only and is intended for fresh preview/staging copies. It does not mutate production and it rejects an existing `OUTPATIENT_SCHEDULE` sheet whose headers do not exactly match the authoritative schema.

## Confirmed DAILY_INPUT reader API mismatch

The C diagnostic log reported:

```text
phase=python preparation
patch_index=0
sheet=
cell=
error_number=2
error={
  "ok": false,
  "error": "'DailyInputReadResult' object has no attribute 'source_changes'"
}
```

This proves the failure occurred during Python preparation before any in-workbook patch was applied. The in-session VBA shell therefore failed closed correctly.

Source comparison showed that the isolated recovery branch still had the older `DailyInputReadResult` contract, while the recovered in-place preparer expects the later reversible-input reader contract. The later contract adds:

- `DailyInputSourceChange`,
- `DailyInputReadResult.source_changes`,
- `ignore_rows`,
- `allow_empty`.

These fields/options are required by the in-session / daily-change-ledger flow and are not optional cosmetic additions.

Bounded recovery fixes:

- `6212c10` recover the complete DAILY_INPUT reader source-change contract,
- `4871371` add focused regressions for `source_changes`, `ignore_rows`, and `allow_empty`.

No workbook persistence code was added by this reader recovery. The reader remains openpyxl/read-only with respect to source workbook persistence.

## Next exact step

1. Build a fresh uniquely named `UNIFIED_CLOSE_DAY_RECOVERY_20261030_D.xlsm` from immutable `Rehab_Center_System_v27_1.xlsm` without `--overwrite`.
2. Verify source SHA unchanged and the strengthened application contract.
3. Read-only verify D contains `OUTPATIENT_SCHEDULE` with exact headers before native smoke.
4. Native-test outpatient manager and DAILY_INPUT on D.
5. If either fails, freeze D as evidence and diagnose from exact native/log output; do not repair D in place.
6. Only if both are healthy may `CloseOperationalDay` be tested.
