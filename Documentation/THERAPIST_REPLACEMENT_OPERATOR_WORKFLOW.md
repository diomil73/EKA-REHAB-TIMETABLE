# Therapist Replacement Operator Workflow

## Status

Locked design decision after native V15 therapist-absence testing.

## Core rule

No therapist-absence replacement may be committed automatically merely because the solver found a feasible or globally optimal assignment.

The solver is an advisory engine. The operator is the decision maker.

For every displaced/orphan patient session, the system must require an explicit operator choice before an ACTIVE replacement fact is created.

## Required workflow

`therapist absence -> identify displaced patient sessions -> compute valid options -> show one patient at a time -> operator chooses -> validate chosen option against all rules -> explicit confirmation -> persist replacement -> recompute remaining patients -> next patient`

The system must never jump directly from therapist absence to persisted global assignments.

## Per-patient screen

The guided replacement screen must present one displaced patient at a time. It must show, in operator language:

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
2. ask the system to select the best valid option according to the locked replacement criteria.

Even when the operator asks for the automatic recommendation, the resulting choice must still be shown before it is committed.

## Double-booking

Double-booking is not a separate primary workflow. It is an exceptional option inside the same per-patient replacement decision when no normal option is preferable or available.

It must be clearly labelled and require explicit confirmation.

## Suspensions

Treatment suspension is a last-resort operator-visible outcome after valid replacement options have been exhausted or explicitly rejected. It must not be created silently by the solver.

## UI requirements

The old dense `frmDoubleBookingReview` list is not an acceptable main therapist-replacement UI.

The replacement UI must be calm and sequential. One patient is the focal item. The operator must not need to interpret internal session IDs, raw solver rows, TSV-like tables, or technical warning strings.

Recommended interaction hierarchy:

- large patient header;
- original appointment summary;
- 3-5 ranked choice cards/rows;
- `Καλύτερη αυτόματη πρόταση` action;
- `Επιλογή` on each option;
- dedicated confirmation area;
- `Επόμενος ασθενής` only after the current patient is resolved;
- progress indicator such as `Ασθενής 2 από 7`.

## Persistence rule

`DAILY_REPLACEMENT_FACTS` receives an ACTIVE row only after operator confirmation of that patient assignment.

The remaining unresolved queue is recomputed after every accepted assignment so later choices respect earlier operator decisions and all capacity/conflict rules.

## V15 finding

V15 proved that the current implementation violates this contract in two ways:

1. `prepare_global_replacements_in_place.py` can return `mode = resolved` and emit replacement rows for automatic persistence;
2. the current guided UI is a dense double-booking review table rather than the required per-patient operator workflow.

Both behaviours are now considered defects, not intended functionality.
