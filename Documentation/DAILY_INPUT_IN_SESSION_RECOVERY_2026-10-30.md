# DAILY_INPUT In-Session Recovery Checkpoint — 2026-10-30

## Trigger

Native smoke on `UNIFIED_CLOSE_DAY_RECOVERY_20261030_B.xlsm` exposed two independent problems before close-day testing:

1. The outpatient menu reported that the outpatient schedule form was unavailable.
2. DAILY_INPUT accepted a patient absence but later failed with `Close of object '_Workbook' failed`.

The B workbook is frozen as evidence. It must not be overwritten or repaired in place.

## Diagnosis

The outpatient menu wrapper caught every error from `frmOutpatientSchedule.Show` and converted it into the same misleading "form not installed" message. The outpatient form source itself matched the later established source, so the immediate recovery need was diagnostic visibility rather than form replacement.

The recovery branch still contained the old DAILY_INPUT authoritative preview/worker flow. That path created `DAILY_INPUT_APPLIED_PREVIEW.xlsm`, launched `authoritative_commit_worker_cli.py`, and then called `ThisWorkbook.Close SaveChanges:=True`. This directly matched the native smoke failure and violated the established in-session mutation architecture.

## Source-only dependency recovery

The established read-only DAILY_INPUT/global-replacement preparers were recovered first. Collection then exposed bounded missing dependencies, all recovered selectively and without workbook mutation:

- `rehab_core.reclining_adjacency`
- `build_service_id_by_session` structured special-service mapping
- pure `daily_suspension_resolution` conversion API

The COM persistence half of the later `daily_suspension_resolution.py` was intentionally not recovered because it opened/closed a workbook. Only the pure date-specific suspension conversion API was retained.

Operator-confirmed source-only gate after bounded dependency recovery:

`38 passed in 3.85s`

## In-session VBA recovery

Recovered established in-session DAILY_INPUT action:

- prepares a validated TSV payload using `prepare_daily_input_in_place.py`,
- saves the currently open workbook before validation with events suppressed,
- applies patches natively in the same workbook,
- semantically reads back each patch,
- creates a rollback copy before mutation,
- restores from rollback on write/readback failure,
- refreshes therapist workload and MASTER views,
- routes therapist-absence cases to the in-session global replacement workflow,
- never calls `ThisWorkbook.Close`, `Application.OnTime`, or `authoritative_commit_worker_cli.py`.

Recovered thin global replacement VBA action:

- uses `prepare_global_replacements_in_place.py`,
- requires explicit operator confirmation,
- persists replacement facts and treatment suspensions in-session,
- applies projection patches through the validated DAILY_INPUT patch function,
- uses rollback and diagnostics,
- preserves explicit guided double-booking review,
- never uses close/reopen worker plumbing.

## Application contract strengthening

`modGlobalReplacementAction` is now a required application component because the recovered `modDailyInputAction` installs and calls it as an operational dependency. Canonical registration build verification now confirms it before saving.

## Outpatient diagnostic recovery

The canonical registration build now patches only `frmRegistrationMenu.cmdOutpatientSchedule_Click` after menu installation. The outpatient form and scheduling engine are unchanged.

On failure, the fresh smoke workbook will display the real `Err.Number` and `Err.Description` from `frmOutpatientSchedule.Show` instead of falsely reporting that the form is not installed. The patch is semantically verified in the VBA CodeModule before save.

## Regression gates added

- DAILY_INPUT action tests reject close/reopen worker plumbing.
- Global replacement action tests require in-session transaction/rollback behavior.
- Production-action safety test rejects `ThisWorkbook.Close`, `Application.OnTime`, authoritative worker plumbing, and console-backed `.Exec` paths for recovered production actions.
- Application contract test requires `modGlobalReplacementAction`.
- Canonical registration build tests require outpatient diagnostic patching and global-action verification.

## Commits in this slice

- `aff9bce` recover in-session DAILY_INPUT VBA action
- `f62c889` recover in-session global replacement VBA action
- `81849f4` replace old close/reopen DAILY_INPUT tests with in-session contract
- `346c828` recover global replacement VBA regression tests
- `14ca274` add production in-session VBA safety gate
- `864889d` require global replacement action in application contract
- `d4ee8d4` lock global replacement action into contract tests
- `d9d21ac` add outpatient runtime diagnostics and verify global action in canonical build
- `fb24a4e` add outpatient diagnostic/build regression tests

## Safety boundary

- Production workbook remains immutable.
- B smoke workbook remains frozen evidence.
- No new `.xlsm` should be built until the new source-only VBA/build gate is green.
- Quarantined manual replacement V2 patch/normalizer tooling remains excluded.
- No bulk merge from `feature/outpatient-scheduling` is permitted.

## Next gate

Run the focused in-session VBA, preparer, application-contract, unified-builder and diagnostic regression suite. Only after it is green may a fresh uniquely named `UNIFIED_CLOSE_DAY_RECOVERY_20261030_C.xlsm` be built for native smoke.
