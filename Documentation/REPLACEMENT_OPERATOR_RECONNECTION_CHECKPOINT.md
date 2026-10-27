# Replacement operator reconnection checkpoint

Date: 2026-10-27

## Purpose

This checkpoint reconnects the already-established therapist-replacement subsystems without changing the protected replacement solver or business rules.

## Protected components

The following remain unchanged and authoritative:

- replacement feasibility and provider policy;
- patient cross-specialty conflict checks;
- provider and student capacity rules;
- service hours;
- reclining-treatment adjacency;
- sequential recalculation after a confirmed manual choice;
- deterministic global replacement planner and unique-slot protection;
- authoritative replacement facts and restore-by-ReplacementId lifecycle.

## Operator routing

When a therapist absence creates displaced sessions, the Excel operator must choose before placement is committed:

- **Yes**: calculate/use the existing global best solution for all unresolved sessions, with explicit confirmation before persistence;
- **No**: resolve patients one-by-one through the manual operator form;
- **Cancel**: stop without creating another replacement.

The DAILY_INPUT route is patched only by exact known markers and fails closed if the expected continuation point is not present.

## Manual path

The manual path is sequential:

1. prepare one unresolved patient using the existing feasibility/policy/service-hour pipeline;
2. show valid therapist/time choices;
3. require explicit operator confirmation;
4. revalidate the selected therapist/time against the current workbook state;
5. create one ACTIVE authoritative replacement fact;
6. rebuild operational projections;
7. synchronize the visible replacement ledger from authoritative facts;
8. recompute the unresolved queue before presenting the next patient.

The preparation and validation Python tools are read-only. They do not save or mutate the workbook.

## DAILY_CHANGE_LEDGER contract

`DAILY_CHANGE_LEDGER` remains the existing visible change/restore ledger beneath `THERAPIST_DAILY`. No second ledger is introduced.

The ledger is now treated as a projection of authoritative `DAILY_REPLACEMENT_FACTS`:

- `ChangeID` is the exact `ReplacementId`;
- patient, original therapist/time, effective therapist/time, reason and status are reconstructed from authoritative workbook data;
- ACTIVE and RESTORED history is preserved in the visible projection;
- the existing `RefreshDailyChangeLedgerView` renders `Από`, `Προς`, `Επαναφορά`, `Επαναφορά επιλεγμένων` and `Επαναφορά όλων`;
- the existing authoritative restore backend continues to restore by `ReplacementId` and full recomputation, not by cell snapshots.

Legacy snapshot columns remain part of the established sheet schema for compatibility/audit context, but they are not the source of truth for restore.

## DAILY_INPUT interaction contract

The operational absence/cancellation entry point is the existing button on `MASTER`, which opens the daily-state input sheet. It is not a registration-menu workflow.

The intended editability contract is now explicit:

- unapplied input rows remain editable and clearable;
- therapist absence `Αιτία` has a native dropdown restored from the established reason list;
- once a DAILY_INPUT change is applied and represented by an ACTIVE `DAILY_CHANGE_LEDGER` record, its source row becomes read-only;
- an applied ACTIVE row must not allow changing therapist/patient, time, status/reason or notes in place;
- reversal of an applied change is performed only through `THERAPIST_DAILY -> Επαναφορά`, so dependent replacements/suspensions are restored transactionally;
- after authoritative restore, the source row is cleared and becomes available again.

The active-row lock is derived only from existing ACTIVE ledger metadata `DAILY_INPUT:*:<source_row>`. No new source of truth is introduced.

## Fail-closed wiring

The installer uses exact transforms to add ledger synchronization after:

- every confirmed manual replacement;
- completion of the selected replacement workflow, including the automatic-all route.

If expected VBA markers are absent or duplicated, the build must stop rather than apply a speculative patch.

## Verification completed on 2026-10-27

Static/regression gates completed:

- focused replacement operator gate: `12 passed in 3.77s`;
- expanded replacement/restore gate: `48 passed in 2.83s`.

Disposable preview build completed successfully from `Excel/previews/RERESOLVED_V90.xlsm`:

- preview: `Excel/previews/MANUAL_REPLACEMENT_SMOKE.xlsm`;
- source unchanged;
- VBA project present;
- registration menu module/form present;
- replacement/review/workload/ledger modules present.

Native smoke findings:

1. therapist absence flow initially failed before replacement routing with message that replacements were not applied;
2. therapist `Αιτία` dropdown was missing natively;
3. daily-state rows were not reliably editable/clearable before repair;
4. user confirmed absence workflow is entered from the MASTER button, not ShowRegistrationMenu or DAILY_INPUT as a standalone menu action;
5. diagnostic log at `2026-10-27 23:46:12` showed current blocker was application-contract validation, not the replacement solver: missing `modOutpatientScheduleManager` and `frmOutpatientScheduleManager`;
6. older `persist suspensions` diagnostic at `23:28:52` is historical and is not treated as the current blocker;
7. application contract repair was run successfully on `MANUAL_REPLACEMENT_SMOKE.xlsm` and verified full contract with no components missing;
8. interaction repair restored the therapist-reason dropdown and preserved existing entries;
9. `Legacy direct-global controls rewired: 0`, so the earlier replacement failure was not caused by a sheet control directly invoking the old global macro;
10. a reliability gap was identified: the first interaction repair unlocked the entire input blocks, which meant an already-applied absence row could still be edited even though it could not be safely deleted;
11. updated interaction repair was run successfully on the same disposable preview and reported:
    - existing entries preserved;
    - therapist reason dropdown verified;
    - unapplied input rows editable/clearable verified;
    - `ACTIVE daily-input rows locked: 0`;
    - `Legacy direct-global controls rewired: 0`;
12. after application-contract repair, the native flow advanced successfully through therapist absence, manual routing, candidate display and explicit `Επιβεβαίωση επιλογής`;
13. immediately after confirmation, Excel reported `Η επιλογή δεν εφαρμόστηκε και αναιρέθηκε.`;
14. source inspection shows that message can only occur after the operator choice has revalidated and an ACTIVE replacement fact has been written temporarily, when `ApplyOperatorProjectionRefresh()` returns false; the replacement fact is then rolled back by design;
15. therefore the current native blocker is post-choice projection refresh, not candidate generation, ranking, manual form routing or choice validation.

Interpretation of `ACTIVE daily-input rows locked: 0`: there were no ACTIVE `DAILY_INPUT:*:<source_row>` ledger records at the time the standalone repair ran. This did not prove live applied-row locking.

## Current code changes after native findings

The DAILY_INPUT interaction repair follows the authoritative editability rule:

- unlock known input ranges first;
- read ACTIVE DAILY_INPUT source rows from `DAILY_CHANGE_LEDGER` status column 12 and metadata column 18;
- re-lock exactly those ACTIVE source rows;
- enable sheet protection so the lock has native UI effect;
- leave unapplied rows editable;
- preserve the reason dropdown and existing entries.

A further native-session protection issue was identified: Excel does not persist `UserInterfaceOnly=True` across workbook close/reopen. A sheet can remain protected for the user while VBA loses permission to write locked cells in the next session. This is a strong candidate for the post-choice projection rollback observed above.

A dedicated `modDailyInputLiveProtection` module and repair path have now been added. It:

- recalculates ACTIVE source-row locks from `DAILY_CHANGE_LEDGER`;
- reapplies `UserInterfaceOnly=True` in the live Excel session;
- hooks the DAILY_INPUT ledger sync so newly ACTIVE rows are locked immediately;
- hooks the manual replacement projection path before `ApplyValidatedDailyInputPayload`;
- writes `%TEMP%\\eka_daily_input_protection_last_error.log` and fails closed if protection refresh cannot be established;
- does not modify replacement feasibility, ranking, validation or the protected global solver.

## Native verification still required

No native Excel acceptance is claimed yet.

Next sequence:

1. run the focused live-protection regression test;
2. install the live-protection hooks into the same disposable `MANUAL_REPLACEMENT_SMOKE.xlsm` using the repair tool with Excel closed;
3. reopen the same preview and repeat the already-reached path: MASTER absence -> apply -> choose manual -> select candidate -> `Επιβεβαίωση επιλογής`;
4. confirm the selection now persists and the projection refresh completes;
5. confirm the applied DAILY_INPUT source row cannot be edited;
6. confirm `THERAPIST_DAILY` shows the visible replacement ledger row with `Από`, `Προς`, `Επαναφορά`;
7. confirm the next patient choices are recomputed;
8. then test authoritative restore.

If the selection still rolls back, read the newest `eka_daily_input_last_error.log`, `eka_replacement_operator_last_error.log` and `eka_daily_input_protection_last_error.log` before any further code change. Only after the manual smoke is green should the automatic-all route be smoke-tested.
