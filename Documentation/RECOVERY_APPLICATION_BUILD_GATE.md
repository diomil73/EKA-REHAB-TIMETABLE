# Recovery Application Build Gate

Last updated: 2026-10-30

## Purpose

This checkpoint records the recovery-only application build safety work performed after returning to the trusted pre-V2 lineage. It does not declare the later replacement V2 experiments safe and it does not modify any production or forensic workbook.

## Safety boundary

- Production workbook remains immutable.
- Forensic `.xlsm` snapshots remain read-only evidence and are not repair targets.
- Recovery source work is isolated on `work/history-statistics-rebuild`.
- Every native workbook test must create a fresh uniquely named copy.
- No existing smoke/preview workbook is overwritten as an experimentation surface.

## Green gates confirmed by operator

### Structural recovery gate

`11 passed in 3.23s`

### Strengthened canonical application gate

`74 passed in 4.60s`

### Pure finalization/history core gate

`21 passed in 1.51s`

### Finalized persistence/history store gate

`32 passed in 4.29s`

### Append-only historical event gate

`40 passed in 2.81s`

### Close-day orchestration gate

`46 passed in 4.88s`

### In-place close-day preparation gate

`49 passed in 4.56s`

### Close-day VBA integration gate

`64 passed in 6.21s`

### Final pre-smoke gate

`64 passed in 5.69s`

This final source-only gate re-ran the strengthened application contract with `modCloseDayAction`, unified builder tests, close-day VBA/in-session preparation, preview/CLI history chain, finalized snapshot/event stores, therapist/day history and daily-state behavior.

### Workload-strip regression gate

`16 passed in 2.86s`

This focused gate covered the canonical workload-strip installation hook together with close-day action, application-contract and unified-builder regression tests.

## First native unified-build attempt

The first fresh native build target was `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_A.xlsm`.

The build safely stopped at the `DAILY_INPUT` stage with:

`Workbook is not a complete application artifact; missing VBA component(s): modTherapistWorkloadStrip`

Important safety result:

- source SHA before/after was unchanged,
- the requested final output file did not exist after the failure,
- therefore the unified builder fail-closed path worked as intended.

Diagnosis: `modTherapistWorkloadStrip` was already an application-contract requirement and its installer already existed, but the current canonical registration build stage did not explicitly install it before the DAILY_INPUT stage asserted the application contract. Static/source tests did not expose this because they did not execute the full native COM build chain.

Recovery fix:

- canonical `build_registration_menu_preview.py` now installs the existing therapist workload strip in the same explicit operational-module installation session as `modCloseDayAction`,
- the pre-clean stage removes any stale workload-strip module before rebuilding,
- the native stage verifies `modTherapistWorkloadStrip` exists before saving,
- a dedicated regression test now locks this canonical build hook.

Commits:

- `51c6a23` install/verify workload strip in canonical registration build
- `2db84df` regression test for workload strip canonical build hook

## Successful unified B smoke candidate

A fresh second build target, `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm`, was created successfully after the workload-strip fix.

Operator verification recorded:

- source SHA remained unchanged,
- output file exists and was not overwritten,
- output size: `382507` bytes,
- output last-write time: `30/10/2026 11:14:07 μμ`,
- `DAILY_INPUT` is present,
- `DAILY_INPUT!B2` reads `30/10/2026`,
- output SHA256 begins `92AC14E820A35F8E22EFF282880BB640C1AE360359B...` as reported by PowerShell.

The B workbook is now the only native smoke candidate for the next step. Do not rebuild or overwrite it. Any further experiment must use a new uniquely named copy.

## Recovered 2026-10-24 roadmap alignment

The roadmap sequence remains:

`Scheduling engine -> Day closure/history -> Statistics/reporting -> Patient progress -> App UI -> Integrated beta -> Hardening -> Production`.

The old effort estimate (`160h baseline / 190h safe budget`) is retained only as historical planning context, not as a current remaining-effort metric.

## Structural recovery fixes now recorded

The canonical outpatient form stage installs the established outpatient schedule manager instead of depending on prior workbook state.

The unified builder verifies the application contract after the full build and rejects/deletes structurally incomplete output.

The application contract now covers the explicit canonical stack including:

- `modApplicationShell`
- `modMasterToolbar`
- `modRegistrationMenu`
- `frmRegistrationMenu`
- `frmNewPatient`
- `frmNewTherapist`
- `frmNewStudent`
- `modPatientRegistrationBridge`
- `frmEditPatient`
- `modDailyInputAction`
- `modTherapistWorkloadStrip`
- `modOutpatientSchedule`
- `frmOutpatientSchedule`
- `modOutpatientInSession`
- `modSchedulingConflict`
- `frmSchedulingConflict`
- `modOutpatientScheduleEdit`
- `modOutpatientScheduleManager`
- `frmOutpatientScheduleManager`
- `modCloseDayAction`

## Selective history recovery completed so far

Recovered core, persistence, append-only event, close-day orchestration, in-session preparation and VBA integration remain isolated on the recovery branch and are individually test-gated.

## Close-day safety properties now recovered

- finalization fails closed when unresolved sessions remain,
- historical snapshots reject unresolved state,
- therapist/day rows distinguish occupied timeslots from active sessions,
- append-only day events use stable IDs and timezone-aware timestamps,
- native persistence uses semantic read-back,
- in-session preparation is read-only and verifies source SHA unchanged,
- VBA writes `SNAPSHOT`, `THERAPIST` and `EVENT` blocks into the currently open workbook,
- VBA preflights duplicate/schema state before writing,
- every written row is semantically read back,
- failures roll back created rows/sheets,
- no close/reopen handoff is used,
- technical history sheets are protected and `VeryHidden`,
- duplicate close attempts are handled as an already-finalized state.

## Important exclusion

Do not reintroduce quarantined replacement V2 patch/repair tooling. Do not bulk merge the later feature branch. Recovery remains selective and test-gated.

## Next exact step

1. Open only `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` for native Excel smoke.
2. Verify application/menu shell and outpatient flow without modifying production or forensic workbooks.
3. Verify DAILY_INPUT native interaction on the B copy.
4. Run `CloseOperationalDay` only after the preceding smoke checks are healthy.
5. Record the native close-day result before moving to amendment/version-chain work or the next roadmap vertical.
