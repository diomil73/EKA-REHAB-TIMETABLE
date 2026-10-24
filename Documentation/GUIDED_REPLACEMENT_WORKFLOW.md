# Guided replacement workflow

## Current production contract

Replacement handling is one daily workflow, not separate therapist/student programs.

1. Enter all known therapist absences for the concrete date.
2. Apply cancellations/no-shows and surviving ACTIVE replacement facts.
3. Solve all remaining displaced sessions together with the global replacement solver.
4. Preserve hard feasibility first, then optimize scarcity, reclining adjacency, original time, workload balance and the weak 13:00 preference.
5. Keep therapists and registered students in the same solver while preserving provider kind and provider-specific capacity/capability rules.
6. If the normal global solve leaves a session unresolved, evaluate the existing last-resort double-booking review path before final suspension.
7. A double-booking/over-capacity assignment may be persisted only after explicit user confirmation.
8. If no acceptable double-booking option exists, or the operator explicitly declines the exception, finalize the session as `Αναστολή θεραπείας`.
9. Recompute the effective daily state and require unresolved sessions = 0 before final operational output.

The base/recurring schedule remains immutable. Daily replacements and suspensions are overlays/facts.

## Already implemented and verified

The following pieces are already operationally available:

- global scarcity-aware assignment rather than sequential greedy choice;
- protection of scarce viable slots;
- reclining/ΦΘ adjacency preference;
- original-time preference below higher-priority constraints;
- workload balancing;
- weak 13:00 avoidance;
- persistent special pool/robotic operating-hour policy;
- authoritative STUDENTS registry integration and distinct `STUDENT` provider kind;
- explicit user-confirmed double-booking review/apply subsystem;
- explicit unresolved-to-`TREATMENT_SUSPENDED` fallback;
- native suspension persistence/read-back and final daily statistics with unresolved = 0.

## Remaining integration gap before replacement-quality closure

The double-booking subsystem and the suspension subsystem are individually green, but the main production sequencing does not yet join them into one last-resort decision flow.

At present the main global replacement tool can turn solver leftovers directly into `Αναστολή θεραπείας`, while the double-booking review is exposed through its own review/apply UI path. Therefore the production flow can finalize a suspension before offering an otherwise valid user-confirmable last-resort exception.

The final integration slice must enforce this order:

`normal global solve -> unresolved review -> explicit double-booking decision -> remaining unresolved suspension -> final recompute`

Required safety semantics:

- never create a double-booking automatically;
- never infer approval from opening/closing the review form;
- only solver-unresolved sessions are eligible for the exception review;
- confirmed exceptions must be revalidated immediately before persistence;
- the review UI must provide an explicit route to decline the exception and continue to `Αναστολή θεραπείας`;
- if an unresolved session has no valid double-booking option, it may proceed directly to suspension;
- finalization still requires unresolved = 0;
- one authoritative replacement engine remains in production; the review is an exception stage, not a second solver.

## Student presentation note

Active registered students are shown by real name in operational views, with distinct student presentation/workload styling. Numbered `φοιτ N` labels are physical placeholders only and return after the corresponding active placement ends. Student identity remains authoritative through StudentID/provider kind and historical records.

## Closure criterion

The replacement-quality milestone can be declared complete when a native end-to-end smoke proves both branches of the final decision gate:

- an unresolved session with a valid exception option is shown for explicit review and can be confirmed without close/reopen regression;
- a declined or unavailable exception is finalized as `Αναστολή θεραπείας`;
- the resulting daily state, ledger/statistics and `THERAPIST_DAILY` projection agree and report zero unresolved sessions.
