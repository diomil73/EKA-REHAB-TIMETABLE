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

## Implemented production sequencing gate

`Python/tools/apply_global_daily_replacements.py` now enforces the required ordering in the production path.

After the normal global solve it builds double-booking review items for solver-unresolved sessions before creating any suspension facts.

If at least one unresolved session has a valid review option, the production apply safety-stops before `build_unresolved_suspensions(...)` unless the operator decision is explicit.

Current explicit rejection contract:

- normal run: reviewable unresolved session -> safety stop, no preview/suspension output;
- acceptance: use the existing confirmed-double-booking review/apply path, which revalidates and persists the ACTIVE exception; rerunning the global flow then reads that persisted fact and the session is no longer unresolved;
- rejection: rerun the production apply only after the operator explicitly rejected the offered exception, using `--reject-double-booking`; only then may the remaining unresolved sessions become `Αναστολή θεραπείας`;
- unresolved sessions with no legal double-booking option may proceed directly to suspension because no user-confirmable exception exists.

This is fail-closed. Opening or closing the review UI is never treated as approval or rejection.

## Already implemented and verified

The following pieces are operationally available:

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
- native suspension persistence/read-back and final daily statistics with unresolved = 0;
- production sequencing guard that evaluates double-booking review before suspension fallback.

Regression coverage in `Tests/test_apply_global_daily_replacements.py` now guards the source ordering:

`build_double_booking_review_items -> explicit decision gate -> build_unresolved_suspensions`

## Remaining native closure work

The backend sequencing gap is closed. The remaining closure evidence is native end-to-end verification of the joined user flow.

Required native smoke:

1. prepare a date where the normal global solver leaves at least one unresolved session with a valid double-booking option;
2. run the normal production apply and verify it safety-stops before creating suspension output;
3. confirm one offered exception through the existing review UI and verify ACTIVE fact + immediate `THERAPIST_DAILY` projection without close/reopen regression;
4. rerun the global flow and verify the accepted session is no longer unresolved;
5. in an independent scenario, explicitly reject an offered exception, rerun with the explicit rejection path, and verify `Αναστολή θεραπείας` persistence;
6. verify final daily state, ledger/statistics and `THERAPIST_DAILY` agree with unresolved = 0.

No separate replacement engine should be introduced for this smoke or for final UI polish.

## Student presentation note

Active registered students are shown by real name in operational views, with distinct student presentation/workload styling. Numbered `φοιτ N` labels are physical placeholders only and return after the corresponding active placement ends. Student identity remains authoritative through StudentID/provider kind and historical records.

## Closure criterion

The replacement-quality milestone can be declared complete when the native end-to-end smoke proves both branches of the final decision gate:

- an unresolved session with a valid exception option is shown for explicit review and can be confirmed without close/reopen regression;
- a declined or unavailable exception is finalized as `Αναστολή θεραπείας`;
- the resulting daily state, ledger/statistics and `THERAPIST_DAILY` projection agree and report zero unresolved sessions.
