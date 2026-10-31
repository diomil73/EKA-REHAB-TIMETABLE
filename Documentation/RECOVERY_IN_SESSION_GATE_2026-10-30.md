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

### Operator visibility UX gate

`34 passed in 3.09s`

This focused gate confirms the presentation-only follow-up derived from the successful native D smoke:

- successful automatic replacements expose a current-action summary containing patient, replacement therapist and replacement time,
- summary generation is scoped to only the replacement fact rows appended by the current action,
- normal DAILY_INPUT refresh no longer auto-renders therapist workload badges on `MASTER_SCHEDULE`,
- the existing workload button remains available for explicit operator request,
- production actions remain in-session,
- canonical builder installs the visibility patch after workload/global replacement modules are present.

The solver/global assignment engine, scoring, capacity rules, double-booking review and persistence semantics remain unchanged.

### COM-safe operator visibility gate

`34 passed in 2.58s`

This repeat gate confirms the fix for the failed E build verification:

- post-install summary verification no longer depends on the Unicode arrow glyph surviving Excel/VBA COM code insertion unchanged,
- the summary presentation uses ASCII `->`,
- post-install verification requires the semantic tokens `PatientDisplayName(patientId)`, `replacementTherapist`, and `replacementTime`,
- workload visibility and all in-session production safety regressions remain green.

### Restored operator + MASTER daily overlay gate

`67 passed, 1 warning in 3.21s`

This source gate confirms the recovered behavior required after native G exposed missing contracts:

- patient absence/cancellation keeps recurring MASTER therapy text visible and applies presentation strike-through,
- MASTER `Κατάσταση` uses the dated non-active reason rather than leaving the patient marked present,
- the MASTER writeback contract remains narrow: therapy overlays only in `D:J`, daily patient status only in `M2:M5000`, with identity columns and `K:L` still protected,
- restored replacement operator routes DAILY_INPUT through global `ΝΑΙ` / manual `ΟΧΙ` / cancel,
- manual mode exposes replacement, same-day postponement, next, previous and automatic remainder controls,
- therapist absence reason is surfaced to the operator,
- manual choices are revalidated against the established replacement rules,
- automatic remainder reuses the unchanged global solver,
- the application contract requires 23 VBA components including `modReplacementOperator` and `frmReplacementDecision`,
- production actions remain in-session and no automatic workload UI is published on MASTER.

The one warning is the already known Python 3.14/openpyxl `ZipFile.__del__` cleanup warning and is not a workbook-mutation signal.

## Safety boundary

- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` remains frozen native-smoke evidence and must not be overwritten or repaired in place.
- Production workbook remains immutable.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm` is frozen as native-smoke evidence and must not be repaired in place.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261030_D.xlsm` is frozen as successful native evidence and must not be repaired in place.
- The attempted `UNIFIED_CLOSE_DAY_RECOVERY_20261031_E.xlsm` build failed before artifact creation; no E workbook exists and no retry should reuse that name.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261031_F.xlsm` and `UNIFIED_CLOSE_DAY_RECOVERY_20261031_G.xlsm` remain frozen recovery evidence and must not be repaired in place.
- `UNIFIED_CLOSE_DAY_RECOVERY_20261031_H.xlsm` is the current fresh 23-component native candidate. It must not be overwritten or repaired in place if native testing exposes a defect.
- Close-day native testing remains blocked until patient-absence MASTER rendering/status and restored therapist-replacement operator behavior are both verified natively on H.

## Fresh H unified build

Operator-confirmed canonical build of `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_H.xlsm` succeeded after the restored MASTER daily overlay, narrow MASTER status write zone, replacement operator recovery and 23-component application contract changes.

Builder output confirmed every unified stage:

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
- application contract count: `23`.

Artifact metadata reported by the operator:

- path: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_H.xlsm`,
- size: `450592` bytes,
- last write: `31/10/2026 5:09:02 μμ`,
- SHA256 begins `92FB49EF6ACC61A119368A6008EF43A260298E28ADB...`.

H is therefore the first native candidate that carries the restored 23-component replacement-operator application contract together with the narrow MASTER daily status write zone.

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

## Fresh D unified build

Operator-confirmed canonical build of `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_D.xlsm` succeeded after the outpatient-sheet and DAILY_INPUT-reader fixes.

Builder output again confirmed all unified stages and reported:

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
- `OUTPATIENT_SCHEDULE PRESENT: True`,
- exact headers: `PatientID`, `Ασθενής`, `Θεραπεία`, `Ώρα`, `Ημέρες`, `Θεραπευτής`.

Artifact metadata reported by the operator:

- path: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_D.xlsm`,
- size: `425218` bytes,
- last write: `31/10/2026 12:18:17 πμ`,
- SHA256 begins `3C3F51B808F840376CF08EC29B5111BC5F890812F2F...`.

D is therefore the first candidate that satisfies both the 21-component VBA application contract and the authoritative outpatient source-sheet contract before native smoke.

## Native D smoke findings

Operator-confirmed D smoke results:

1. Outpatient schedule management opened and its controls worked correctly. The initialized outpatient source contained no registered outpatient patient, so there was no real patient available to create a schedule for; this is data state, not a UI/runtime failure.
2. DAILY_INPUT with one patient absence completed successfully in-session and returned to `MASTER_SCHEDULE` correctly.
3. The MASTER refresh displayed therapist workload information after the operation.

These results close the two native failures found on B/C:

- outpatient manager is now operational,
- DAILY_INPUT no longer closes/reopens the workbook and no longer fails on the reversible-reader contract.

Two UX gaps remain and were explicitly reconfirmed by the operator:

- automatic replacement is not transparent enough: after success the operator must be told which patient was assigned to which replacement therapist and at what time,
- therapist workload must not be published automatically across the MASTER view after normal DAILY_INPUT/replacement operations; workload information should be available only when explicitly requested.

## Operator visibility UX patch

The solver/global assignment engine remains unchanged.

Bounded presentation-only fixes:

- `79fb796` adds `operator_visibility_patch_vba.py`, which patches the installed normal-production VBA so DAILY_INPUT refresh stops auto-rendering workload badges and successful global replacement shows a summary of only the replacement fact rows appended by the current action,
- replacement summary resolves `PatientID` through `PATIENTS` and displays patient, replacement therapist and replacement time,
- the existing `Φόρτος θεραπευτών` button remains available for explicit operator request,
- `adb9309` installs the visibility patch in the canonical registration/operational build session after workload/global replacement modules are present,
- `c99eb0a` adds regressions for replacement summary, current-action row scoping, optional workload visibility, and canonical build installation.

This patch does not change assignment scoring, candidate generation, capacity rules, double-booking review, suspension semantics, or replacement persistence.

## Failed E build and COM-safe verification fix

The first fresh E build attempt stopped fail-closed during registration/operational module installation with:

```text
SAFETY STOP: Operational module installation failed during patching operator visibility: Replacement summary details were not installed
```
