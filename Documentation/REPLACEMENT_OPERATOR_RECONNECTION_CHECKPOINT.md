# Replacement operator reconnection checkpoint

Date: 2026-10-29

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

Manual preparation must remain lightweight. It must not invoke the full global assignment solver merely to decorate one displayed choice as a recommendation. The automatic **Yes** route remains the owner of the authoritative global solve. Manual mode still uses the established feasibility, policy, capacity, service-hour and queue logic to enumerate valid choices.

## DAILY_CHANGE_LEDGER contract

`DAILY_CHANGE_LEDGER` remains the existing visible change/restore ledger beneath `THERAPIST_DAILY`. No second ledger is introduced.

The ledger is treated as a projection of authoritative `DAILY_REPLACEMENT_FACTS`:

- `ChangeID` is the exact `ReplacementId`;
- patient, original therapist/time, effective therapist/time, reason and status are reconstructed from authoritative workbook data;
- ACTIVE and RESTORED history is preserved in the visible projection;
- the existing `RefreshDailyChangeLedgerView` renders `Από`, `Προς`, `Επαναφορά`, `Επαναφορά επιλεγμένων` and `Επαναφορά όλων`;
- the existing authoritative restore backend continues to restore by `ReplacementId` and full recomputation, not by cell snapshots.

Legacy snapshot columns remain part of the established sheet schema for compatibility/audit context, but they are not the source of truth for restore.

## DAILY_INPUT interaction contract

The operational absence/cancellation entry point is the existing button on `MASTER`, which opens the daily-state input sheet. It is not a registration-menu workflow.

The intended editability contract is explicit:

- unapplied input rows remain editable and clearable;
- therapist absence `Αιτία` has a native dropdown restored from the established reason list;
- once a DAILY_INPUT change is applied and represented by an ACTIVE `DAILY_CHANGE_LEDGER` record, its source row becomes read-only;
- an applied ACTIVE row must not allow changing therapist/patient, time, status/reason or notes in place;
- reversal of an applied change is performed only through `THERAPIST_DAILY -> Επαναφορά`, so dependent replacements/suspensions are restored transactionally;
- after authoritative restore, the source row is cleared and becomes available again.

The active-row lock is derived only from existing ACTIVE ledger metadata `DAILY_INPUT:*:<source_row>`. No new source of truth is introduced.

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
5. diagnostic log at `2026-10-27 23:46:12` showed the blocker was application-contract validation, not the replacement solver: missing `modOutpatientScheduleManager` and `frmOutpatientScheduleManager`;
6. application contract repair was run successfully on `MANUAL_REPLACEMENT_SMOKE.xlsm` and verified full contract with no components missing;
7. interaction repair restored the therapist-reason dropdown and preserved existing entries;
8. `Legacy direct-global controls rewired: 0`;
9. updated interaction repair reported existing entries preserved, therapist reason dropdown verified, unapplied input rows editable/clearable verified, `ACTIVE daily-input rows locked: 0`, and no legacy direct-global controls rewired;
10. after application-contract repair, the native flow advanced through therapist absence, manual routing, candidate display and explicit `Επιβεβαίωση επιλογής`;
11. immediately after confirmation, Excel reported `Η επιλογή δεν εφαρμόστηκε και αναιρέθηκε.`;
12. source inspection showed that message occurs after operator choice revalidation and temporary ACTIVE replacement persistence when `ApplyOperatorProjectionRefresh()` returns false, after which the replacement fact is rolled back by design.

Interpretation of `ACTIVE daily-input rows locked: 0`: there were no ACTIVE `DAILY_INPUT:*:<source_row>` ledger records at the time the standalone repair ran. This did not prove live applied-row locking.

## Live-protection repair added after the 2026-10-27 smoke

Excel does not persist `UserInterfaceOnly=True` across workbook close/reopen. A dedicated `modDailyInputLiveProtection` module and repair path were added to:

- recalculate ACTIVE source-row locks from `DAILY_CHANGE_LEDGER`;
- reapply `UserInterfaceOnly=True` in the live Excel session;
- hook DAILY_INPUT ledger sync so newly ACTIVE rows are locked immediately;
- hook the manual replacement projection path before `ApplyValidatedDailyInputPayload`;
- write `%TEMP%\\eka_daily_input_protection_last_error.log` and fail closed if protection refresh cannot be established;
- avoid changing replacement feasibility, ranking, validation or the protected global solver.

## Native diagnosis on 2026-10-29

The same disposable preview was reopened and a therapist absence was applied successfully. The operator router appeared normally and **No / manual** was selected.

The workflow then appeared to hang before the manual candidate form appeared.

Evidence collected while Excel remained open:

- fresh DAILY_INPUT payload/response files were created at `2026-10-29 21:24`, proving the daily-input application stage had completed;
- no new replacement-operator rows/response had been completed yet;
- two Python processes existed for the same `prepare_replacement_operator_decision.py` command, one launcher and one real worker;
- the real worker PID `21032` increased from `381.5625` to `395.15625` CPU seconds across the observation interval, proving active computation rather than file-lock or I/O waiting;
- Excel itself reported `Responding=True` while synchronously waiting for the Python preparation process;
- terminating only those two Python processes caused the VBA fail-safe message `Δεν ήταν δυνατός ο υπολογισμός των επιλογών αναπλήρωσης.` and returned control to Excel;
- the preparation tool is read-only with respect to the workbook, so this diagnostic termination did not create an authoritative replacement fact.

Source inspection identified the cause of the excessive manual preparation cost: `prepare_replacement_operator_decision.py` first computed valid manual options, then invoked the full `build_global_replacement_assignments()` solver only to mark one displayed option as `is_global_recommendation`. This made opening the manual form depend on completion of a full global solve.

This does **not** prove a defect in the protected global solver itself. The regression is architectural: the manual UI path unnecessarily depended on the full automatic solver before it could display valid choices.

## Fix on 2026-10-29

Manual preparation has been decoupled from the full global solve:

- `prepare_replacement_operator_decision.py` still builds the authoritative unresolved queue;
- it still applies existing feasibility, policy, capacity, student, service-hour and cancellation/absence rules;
- it still chooses the next target deterministically and emits valid therapist/time choices;
- it no longer imports or invokes `build_global_replacement_assignments()` or reclining-adjacency scoring merely to decorate a recommendation;
- compatibility fields remain present, with `is_global_recommendation=0` and blank recommended provider/time;
- the automatic **Yes** route remains unchanged and still calls `ResolveDailyReplacementsInPlace()` for the protected global solution.

Commits:

- `a3f4c7ab79d03cb16da1e07e415ac74812d4cb7e` `Keep manual replacement preparation out of global solve`
- `9df56759ee7786b5ccefa2dd7a13c0be35668d29` `Guard lightweight manual replacement preparation`

## Native verification still required

No native Excel acceptance is claimed yet.

Next sequence:

1. pull the branch and run the new lightweight-manual regression guard together with the established replacement operator/live-protection tests;
2. do **not** run another interaction or live-protection repair on the current preview unless a code change specifically requires reinstalling generated VBA;
3. reopen the smoke workbook and repeat: MASTER -> Κατάσταση ημέρας -> therapist absence -> apply -> choose **No/manual**;
4. verify the manual candidate form appears promptly and no long-running `prepare_replacement_operator_decision.py` worker remains;
5. confirm one manual candidate;
6. if confirmation again rolls back with `Η επιλογή δεν εφαρμόστηκε και αναιρέθηκε.`, return to the already-isolated projection/protection diagnostics rather than the solver;
7. only after one manual replacement persists, verify ACTIVE DAILY_INPUT row locking, visible `THERAPIST_DAILY` ledger projection, next-patient recomputation and authoritative restore;
8. only after the manual smoke is green should the automatic-all route be smoke-tested.

After any generic native failure, read the relevant fresh diagnostic log before another retry. Do not repeat blind smokes.
