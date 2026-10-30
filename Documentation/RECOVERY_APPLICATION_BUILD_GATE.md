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

This latest gate covered the recovered close-day VBA action, read-only in-session payload preparation, preview/CLI history chain, application contract and unified builder tests.

## Recovered 2026-10-24 roadmap alignment

A user-provided roadmap snapshot dated 2026-10-24 confirms the intended project sequence after the scheduling engine milestone:

`Scheduling engine -> Day closure/history -> Statistics/reporting -> Patient progress -> App UI -> Integrated beta -> Hardening -> Production`.

The roadmap explicitly places `Κλείσιμο ημέρας + finalized historical snapshot` as the next implementation slice and matches `STATISTICS_HISTORY_DESIGN.md`: fail-closed finalization, versioned snapshots, append-only events, therapist/day history, later patient/department reporting, progress, application UI, beta, hardening and production.

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

The close-day component was added to the contract only after its source, in-session preparation path, canonical build hook and regression tests were all green.

## Selective history recovery completed so far

Recovered pure-core source:

- `Python/rehab_core/day_finalization.py`
- `Python/rehab_core/therapist_day_history.py`
- `Tests/test_day_finalization.py`
- `Tests/test_therapist_day_history.py`

Recovered persistence/store source:

- `Python/rehab_excel/finalized_day_snapshots.py`
- `Python/rehab_excel/therapist_day_history_store.py`
- `Tests/test_finalized_day_snapshots.py`
- `Tests/test_therapist_day_history_store.py`

Recovered append-only event source:

- `Python/rehab_core/day_history_events.py`
- `Python/rehab_excel/day_history_event_store.py`
- `Tests/test_day_history_events.py`
- `Tests/test_day_history_event_store.py`

Recovered close-day orchestration and in-session preparation:

- `Python/tools/build_finalized_day_snapshot_preview.py`
- `Tests/test_build_finalized_day_snapshot_preview.py`
- `Python/tools/close_day.py`
- `Tests/test_close_day_cli.py`
- `Python/tools/prepare_close_day_in_place.py`
- `Tests/test_prepare_close_day_in_place.py`

Recovered VBA integration:

- `Python/rehab_excel/close_day_action_vba.py`
- `Tests/test_close_day_action_vba.py`
- canonical `build_registration_menu_preview.py` explicitly installs and verifies `modCloseDayAction`
- `REQUIRED_APP_COMPONENTS` now requires `modCloseDayAction`

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

## Recovery commits for the close-day integration slice

- `b7f3094` recovered `modCloseDayAction`
- `0f634d0` canonical build installs/verifies close-day action
- `baeae77` close-day VBA/build integration regression test
- `5d7b65b` close-day VBA recovery checkpoint
- `7b60620` require close-day action in application contract
- `c5c9fdb` test close-day action as required application component

## Important exclusion

Do not reintroduce quarantined replacement V2 patch/repair tooling. Do not bulk merge the later feature branch. Recovery remains selective and test-gated.

## Next exact step

1. Run the final pre-smoke gate with the strengthened application contract including `modCloseDayAction`.
2. If green, build one fresh uniquely named unified `.xlsm` from the recovery branch. Do not overwrite any prior smoke/preview workbook.
3. Verify source hash unchanged and final application contract on the new copy.
4. Perform native Excel smoke on the new copy, including outpatient flow, DAILY_INPUT and `CloseOperationalDay`.
5. Only after native proof, treat the recovered close-day/history slice as integrated and move to amendment/version-chain work or the next roadmap vertical.
