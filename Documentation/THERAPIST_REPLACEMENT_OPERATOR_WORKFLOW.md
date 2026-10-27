# Therapist Replacement Operator Workflow

## Status

Locked design decision after native V15 therapist-absence testing.

## Core rule

No therapist-absence replacement may be committed automatically merely because the solver found a feasible or globally optimal assignment.

The solver is an advisory engine. The operator is the decision maker.

For every displaced/orphan patient session, the system must require an explicit operator choice before an ACTIVE replacement fact is created.

## Protected core: do not rewrite during UI/orchestration work

The existing replacement core is already authoritative and must be treated as a protected zone while the V15 operator workflow is repaired.

In particular, do not weaken or replace:

- `rehab_core.global_replacement_planner.plan_global_replacements`;
- replacement feasibility and provider-policy filtering;
- patient cross-specialty conflict protection;
- provider daily-capacity limits;
- service-hours restrictions;
- patient same-time conflict protection;
- provider same-time conflict protection;
- reclining-treatment adjacency penalty;
- original-time preference after higher priorities are satisfied;
- workload balancing;
- the weak preference to avoid 13:00 when an otherwise equivalent solution exists.

The global planner intentionally solves all remaining displaced sessions together. It orders requests with fewer candidates first and maximizes the total number of placed sessions before lower-priority preferences. Therefore a scarce/unique feasible slot is protected from being consumed by another patient who has alternatives.

Regression test `Tests/test_global_replacement_planner.py::test_protects_unique_slot_to_maximize_complete_assignment` is a mandatory guard. Example contract: if patient B can only use 09:00 while patient A can use 09:00, 10:00 or 11:00, B keeps 09:00 and A moves to an alternative.

## Required workflow

The agreed operator workflow combines the earlier sequential manual workflow with the later exact global planner:

`therapist absence -> identify displaced patient sessions -> manual operator decisions, if any -> validate and persist each confirmed manual decision -> recompute operational state after every confirmed manual decision -> global solver evaluates all remaining displaced sessions together -> show resulting proposal -> explicit operator confirmation -> persist confirmed proposal -> continue until no unresolved session remains`

The system must never jump directly from therapist absence to persisted global assignments.

### Manual choices are locks

A manual choice is not merely a UI preference. After the operator confirms a therapist/time for one patient:

1. the assignment is validated against all existing replacement rules;
2. it becomes part of the current operational state;
3. workload, capacity and occupied slots are recalculated;
4. the remaining queue is rebuilt;
5. subsequent manual options and any later global proposal must respect that confirmed assignment.

This preserves the behaviour introduced by the sequential therapist-replacement workflow: rankings cannot be a stale static list calculated before earlier choices.

### Automatic assistance for the remaining patients

If the operator asks for automatic assistance for the unresolved remainder, the existing global solver must be used. It must not be replaced by a greedy per-patient selection.

The solver may recommend the complete remaining plan, but recommendation is not persistence. The proposed assignments must be shown to the operator before ACTIVE replacement facts are created.

If the operator makes no manual choices at all, the same global solver may calculate the full proposal from the original displaced queue. Its scarce-slot/global-completion protections still apply.

## Per-patient screen

The guided replacement screen must present one displaced patient at a time for manual work. It must show, in operator language:

- patient name;
- original therapist;
- original appointment time;
- relevant treatment/service context when needed;
- ranked valid replacement choices;
- therapist name;
- proposed replacement time;
- current therapist occupied-slot load/capacity where relevant;
- any exceptional condition, especially double-booking;
- a clear action for automatic recommendation and a separate explicit confirmation action.

The operator may either:

1. choose a specific valid therapist/time option; or
2. ask the system to calculate the best valid plan for the remaining unresolved patients according to the locked global criteria.

Even when the operator asks for the automatic recommendation, the resulting choice/plan must still be shown before it is committed.

## Double-booking

Double-booking is not a separate primary workflow. It is an exceptional option inside the same replacement decision process when no normal option is preferable or available.

It must be clearly labelled and require explicit confirmation.

The dense legacy `frmDoubleBookingReview` may remain as an implementation detail only until its capability is absorbed into the operator workflow. It must not become the main replacement screen.

## Suspensions

Treatment suspension is a last-resort operator-visible outcome after valid replacement options have been exhausted or explicitly rejected. It must not be created silently by the solver.

## UI requirements

The old dense `frmDoubleBookingReview` list is not an acceptable main therapist-replacement UI.

The replacement UI must be calm and sequential. One patient is the focal item during manual resolution. The operator must not need to interpret internal session IDs, raw solver rows, TSV-like tables, or technical warning strings.

Recommended interaction hierarchy:

- large patient header;
- original appointment summary;
- 3-5 ranked choice cards/rows;
- `Καλύτερη αυτόματη πρόταση για τους υπόλοιπους` action;
- `Επιλογή` on each manual option;
- dedicated confirmation area;
- `Επόμενος ασθενής` only after the current patient is resolved or deliberately deferred to the global proposal;
- progress indicator such as `Ασθενής 2 από 7`;
- separate final proposal summary before bulk persistence of automatically proposed remaining assignments.

## Persistence rule

`DAILY_REPLACEMENT_FACTS` receives an ACTIVE row only after operator confirmation.

For a manual choice, confirmation is per patient. For an automatically calculated remaining plan, the proposal must be displayed and explicitly confirmed before the plan is persisted.

The remaining unresolved queue is recomputed after every accepted manual assignment so later choices respect earlier operator decisions and all capacity/conflict rules.

## Mandatory regression guards before and after UI work

At minimum, keep these behaviours green while changing orchestration/UI:

- unique scarce slot is protected for the patient who has no alternative;
- total number of placed patients has priority over preserving one patient's original time;
- one patient cannot receive two replacements at the same clock time;
- one provider cannot receive two displaced patients in the same clock slot;
- remaining provider capacity is respected globally;
- reclining adjacency outranks original-time preference when both plans place everyone;
- workload balance is used only after higher-priority rules tie;
- 13:00 remains usable rather than leaving a patient unresolved;
- manual accepted assignments update load/availability before the next decision.

## V15 finding

V15 proved that the current implementation violates this contract in two ways:

1. `prepare_global_replacements_in_place.py` can return `mode = resolved` and `modGlobalReplacementAction` currently persists that plan immediately;
2. the current guided UI is a dense double-booking review table rather than the required operator workflow.

These are orchestration/UI defects. They are not a reason to rewrite the global solver or its candidate-policy pipeline.

## 2026-10-27 safety restoration checkpoint

The first repair step is now implemented in the active branch without touching the protected solver or replacement rules.

`modGlobalReplacementAction.ResolveDailyReplacementsInPlace` now requires explicit operator confirmation before `PersistReplacementPayload` can run. Choosing `Όχι` leaves the workbook unchanged and reserves the unresolved sessions for the manual patient-by-patient workflow. The confirmation is specifically labelled as applying the best automatic global solution, so automatic assistance is an explicit operator action rather than the default side effect of therapist absence entry.

Regression coverage in `Tests/test_global_replacement_action_vba.py` asserts that the confirmation gate appears before persistence. This is a safety checkpoint, not the final manual UI. The next implementation step is to connect the existing sequential replacement engine to a calm Excel patient-by-patient interface while keeping this global confirmation gate as the separate automatic-all path.

## V16 experimental overlay withdrawn

The isolated V16 operator overlay was tested natively on 2026-10-27 and is not an accepted replacement baseline. It exposed unsafe orchestration behavior when an already-applied therapist absence was interrupted and then edited to another therapist. The V16 overlay components and their dedicated tests have therefore been removed from the active branch.

This withdrawal does not revert or alter the protected replacement core, the global solver, feasibility rules, adjacency logic, cross-specialty protection, capacity rules, provider policy, service hours, workload balancing, unique-slot protection, or the established double-booking validation path.

Resume from the pre-V16 replacement implementation and treat future work as narrowly scoped orchestration/restore integration. Any future UI change must reuse the existing replacement engine and must pass the established replacement regression gate before native Excel verification.
