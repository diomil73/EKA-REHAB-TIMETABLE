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

On 2026-10-30 the focused recovery gate completed successfully:

`11 passed in 3.23s`

### Strengthened canonical application gate

After expanding the application contract to the full currently-established outpatient application stack, the operator reported:

`74 passed in 4.60s`

This gate covered the application contract, unified builder, patient registration bridge/edit, outpatient form/edit/manager, conflict guard, in-session writer and scheduling conflict popup.

## Structural recovery fixes now recorded

The canonical outpatient form stage now installs the already-established outpatient schedule manager rather than depending on prior workbook state.

The unified builder now verifies the application contract after the full build and rejects/deletes an output that is structurally incomplete.

The application contract now covers components that are explicit outputs of the canonical unified build stages:

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

This intentionally means a workbook that has VBA but has lost the outpatient create form/module cannot be accepted as a complete application artifact.

## Selective history recovery started

The later `feature/outpatient-scheduling` line was compared against the trusted recovery branch. The close-day/history work exists there as a bounded implementation slice, alongside many unrelated later changes and quarantined V2 experiments. Therefore no bulk merge/cherry-pick is allowed.

The first recovered slice is deliberately pure-core only and performs no Excel/COM mutation:

- `Python/rehab_core/day_finalization.py`
- `Python/rehab_core/therapist_day_history.py`
- `Tests/test_day_finalization.py`
- `Tests/test_therapist_day_history.py`

Recovered semantics:

- day finalization is fail-closed when unresolved sessions remain,
- finalized snapshot counts come from effective daily truth,
- therapist/day workload distinguishes occupied timeslots from active sessions,
- replacement workload moves to the effective therapist,
- student replacement work is not misreported as therapist workload,
- non-performed outcomes stay attributable to the original therapist.

No persistence store, close-day CLI, VBA action, or installer hook has been recovered yet. Those remain gated behind the pure-core tests.

## Commits in this recovery line

- `cfbffc3` regression test for manager installation in canonical outpatient form stage
- `b91e57c` safe workbook wrapper for outpatient manager installation
- `ea3b982` canonical outpatient form stage installs manager
- `d4a83ba` unified final application-contract regression tests
- `6ca4d8d` unified builder performs final contract verification
- `542d042` initial recovery build-gate checkpoint
- `5efff1c` strengthened canonical application component contract
- `25e0cb9` regression test for the strengthened component contract
- `9652142` recovered pure day-finalization core
- `3eaa5cb` recovered therapist/day history core
- `b52a087` recovered day-finalization tests
- `d6deb93` recovered therapist/day history tests

## Important exclusion

Do not add close-day/history VBA, replacement-operator V2, or other later operational components to `REQUIRED_APP_COMPONENTS` merely because they existed in later workbook snapshots.

Those components must first be recovered as explicit source/build slices with their own tests and installer path. Do not reintroduce `manual_replacement_operator_v2_patch.py`, `manual_operator_v2_source_normalizer.py`, or related repair tooling.

## Next exact step

1. Run the pure-core finalization/history tests together with the already-green daily-state/statistics tests.
2. If green, recover finalized snapshot persistence and therapist/day history storage as a separate no-VBA slice.
3. Test persistence/version-chain behavior before recovering `close_day.py` or `modCloseDayAction`.
4. Recover the close-day in-session/VBA action only after persistence tests are green.
5. Add `modCloseDayAction` to the application contract only after its canonical installer/build path is explicit and green.
6. Build a fresh uniquely named unified `.xlsm` only after all Python gates are green.
