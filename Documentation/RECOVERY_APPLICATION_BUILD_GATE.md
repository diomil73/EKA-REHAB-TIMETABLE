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

This gate covered the application contract, unified builder, patient registration bridge/edit, outpatient form/edit/manager, conflict guard, in-session writer and scheduling conflict popup.

### Pure finalization/history core gate

`21 passed in 1.51s`

This gate covered recovered day finalization, therapist/day history derivation and existing daily-state behavior. It confirms that the recovered core remains fail-closed for unresolved sessions and preserves the known daily truth semantics.

### Finalized persistence/history store gate

`32 passed in 4.29s`

This gate covered finalized snapshot persistence readers/schema guards, therapist/day history storage, day finalization, therapist/day history derivation and existing daily-state behavior.

### Append-only historical event gate

`40 passed in 2.81s`

This gate covered the append-only event model/store together with the already-green finalized snapshot, therapist/day history and daily-state layers.

### Close-day orchestration gate

`46 passed in 4.88s`

This gate covered the finalized-day preview builder and close-day CLI together with the full already-green history/event stack.

## Recovered 2026-10-24 roadmap alignment

A user-provided roadmap snapshot dated 2026-10-24 confirms the intended project sequence after the scheduling engine milestone:

`Scheduling engine -> Day closure/history -> Statistics/reporting -> Patient progress -> App UI -> Integrated beta -> Hardening -> Production`.

The roadmap explicitly places `Κλείσιμο ημέρας + finalized historical snapshot` as the next implementation slice and specifies the same architecture already recorded in `STATISTICS_HISTORY_DESIGN.md`:

- fail-closed day finalization with unresolved = 0,
- immutable/versioned daily snapshots plus append-only historical events,
- superseded/current-final correction chains,
- patient lifecycle dates and historical patient reporting,
- therapist workload history with occupied timeslots and active sessions as separate metrics,
- department and clinic aggregation from the same finalized facts,
- 15-day patient progress milestones keyed by stable PatientID,
- authoritative holiday/attendance semantics,
- reporting UI with tables first and charts/trends as complementary views,
- MASTER as the eventual application home screen,
- interactive patient/therapist links after historical identity is locked,
- integrated beta, regression, hardening, migration and production packaging only after the historical/reporting layers are stable.

This recovered roadmap is treated as planning evidence and scope continuity, not as permission to trust later experimental implementation. In particular, the statement that the scheduling/replacement engine was complete refers to the known-good engine state at that time and does not validate later quarantined replacement V2 patch/repair work.

The old effort estimate (`160h baseline / 190h safe budget`) is preserved only as historical planning context. It is not used as a current progress metric because substantial parts of the roadmap were implemented after 2026-10-24 and are now being selectively recovered rather than rebuilt.

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

## Selective history recovery

The later `feature/outpatient-scheduling` line was compared against the trusted recovery branch. The close-day/history work exists there as a bounded implementation slice, alongside many unrelated later changes and quarantined V2 experiments. Therefore no bulk merge/cherry-pick is allowed.

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

Recovered append-only historical event layer:

- `Python/rehab_core/day_history_events.py`
- `Python/rehab_excel/day_history_event_store.py`
- `Tests/test_day_history_events.py`
- `Tests/test_day_history_event_store.py`

Recovered close-day orchestration slice:

- `Python/tools/build_finalized_day_snapshot_preview.py`
- `Tests/test_build_finalized_day_snapshot_preview.py`
- `Python/tools/close_day.py`
- `Tests/test_close_day_cli.py`

Recovered read-only in-session preparation slice after the close-day orchestration gate became green:

- `Python/tools/prepare_close_day_in_place.py`
- `Tests/test_prepare_close_day_in_place.py`

The in-session preparer does not mutate the workbook. It validates the current application contract and service date, rejects already-finalized dates, reconstructs the final effective daily truth, creates one validated payload containing atomic `SNAPSHOT`, `THERAPIST` and `EVENT` blocks, and verifies the source workbook SHA is unchanged before returning success.

Persistence/event safety properties recovered:

- finalized snapshot reader validates schema and rejects duplicate `CURRENT_FINAL` rows per service date,
- finalized history rejects unresolved snapshots,
- native writers use Excel COM and semantic read-back rather than openpyxl mutation of operational `.xlsm`,
- therapist/day store rejects duplicate `(SnapshotId, TherapistId)` keys,
- one append batch is constrained to one snapshot/date,
- day-history events use immutable IDs, timezone-aware timestamps and append-only semantics,
- duplicate `EventId` values are rejected both on read and append,
- technical history sheets are protected and `VeryHidden`,
- COM `NumberFormat` assignment is intentionally avoided because it was previously unreliable across Excel builds/locales,
- close-day in-session preparation is read-only for the workbook and deletes/invalidates its payload on validation failure.

No close-day VBA action, amendment workflow or canonical installer hook has been recovered yet. Those remain gated behind the in-session preparation tests.

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
- `2b410b3` recovered finalized snapshot store
- `cec4776` recovered therapist/day history store
- `fda76fe` recovered finalized snapshot store tests
- `5426db3` recovered therapist/day history store tests
- `d1df042` recovered append-only day history event model
- `d86303f` recovered day history event store
- `62a3bc6` recovered day history event tests
- `fad96c9` recovered day history event store tests
- `7c9ed45` recovered finalized-day snapshot preview builder
- `c4daf0a` recovered finalized-day preview regression tests
- `b2a0658` recovered close-day orchestration CLI
- `cc5b2e1` recovered close-day CLI regression tests
- `71ceaad` recovered read-only close-day in-place payload preparation
- `f4d4b84` recovered close-day in-place preparation regression tests

## Important exclusion

Do not add close-day/history VBA, replacement-operator V2, or other later operational components to `REQUIRED_APP_COMPONENTS` merely because they existed in later workbook snapshots.

Those components must first be recovered as explicit source/build slices with their own tests and installer path. Do not reintroduce `manual_replacement_operator_v2_patch.py`, `manual_operator_v2_source_normalizer.py`, or related repair tooling.

## Next exact step

1. Run `test_prepare_close_day_in_place.py` together with the already-green close-day/history stack.
2. If green, recover `close_day_action_vba.py` and its regression tests.
3. Verify the canonical registration/build installer path explicitly installs `modCloseDayAction`.
4. Only then add `modCloseDayAction` to `REQUIRED_APP_COMPONENTS` and its contract tests.
5. Build a fresh uniquely named unified `.xlsm` only after all Python gates are green.
