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

1. Run focused source tests for the new workload-strip canonical build hook plus unified/application contract tests.
2. If green, build a new fresh uniquely named unified preview (`..._B.xlsm`), not the failed `..._A.xlsm` name.
3. Verify source hash unchanged and final application contract on the new copy.
4. Perform native Excel smoke only after the build fully succeeds.
