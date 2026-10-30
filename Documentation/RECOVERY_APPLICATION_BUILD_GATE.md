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

## Green gate confirmed by operator

On 2026-10-30 the focused recovery gate completed successfully:

`11 passed in 3.23s`

Command set:

```powershell
git status --short
git fetch origin
git switch work/history-statistics-rebuild
git pull --ff-only origin work/history-statistics-rebuild

python -m pytest `
  Tests/test_build_outpatient_schedule_form_preview_manager.py `
  Tests/test_unified_builder.py `
  Tests/test_application_contract.py -q
```

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

## Commits in this recovery slice

- `cfbffc3` regression test for manager installation in canonical outpatient form stage
- `b91e57c` safe workbook wrapper for outpatient manager installation
- `ea3b982` canonical outpatient form stage installs manager
- `d4a83ba` unified final application-contract regression tests
- `6ca4d8d` unified builder performs final contract verification
- `542d042` initial recovery build-gate checkpoint
- `5efff1c` strengthened canonical application component contract
- `25e0cb9` regression test for the strengthened component contract

## Important exclusion

Do not yet add close-day/history, replacement-operator V2, or other later operational components to `REQUIRED_APP_COMPONENTS` merely because they existed in later workbook snapshots.

Those components must first be recovered as an explicit source/build slice with their own tests and installer path. The later repository line already contains close-day/history implementation and tests, so the next task is selective recovery, not reimplementation from scratch.

## Next exact step

1. Run the strengthened application-contract focused tests.
2. If green, inspect the later close-day/history implementation and its tests as a self-contained recovery slice.
3. Port only the already-established close-day/history source required by the canonical application path.
4. Add its components to the application contract only after its installer/build stage is explicit and green.
5. Build a fresh uniquely named unified `.xlsm` only after Python gates are green.
