# Guided replacement workflow

## Status

**Replacement-quality milestone: CLOSED end-to-end on 2026-10-24.**

The production decision chain is one workflow for therapists and students:

1. enter all known therapist absences for the concrete date;
2. apply cancellations/no-shows and surviving ACTIVE replacement facts;
3. solve all remaining displaced sessions together with the global replacement solver;
4. preserve hard feasibility first, then scarcity, reclining adjacency, original time, workload balance and weak 13:00 avoidance;
5. keep therapists and active registered students in the same solver while preserving provider kind and provider-specific capability/capacity rules;
6. if the normal global solve leaves a session unresolved, build last-resort double-booking review options before creating any suspension;
7. a double-booking exception may be persisted only after explicit user confirmation;
8. if the user explicitly rejects the offered exception, or no legal exception exists, finalize the remaining orphan as `Αναστολή θεραπείας`;
9. recompute the effective daily state and require `unresolved_sessions == 0`.

The recurring/base schedule remains immutable. Daily replacements and suspensions are authoritative overlays/facts.

## Production sequencing gate

`Python/tools/apply_global_daily_replacements.py` enforces:

`global solve -> review gate -> explicit accept/reject -> remaining suspension -> final recompute`

Fail-closed behavior:

- reviewable unresolved session + no explicit decision => safety stop before suspension;
- explicit ACCEPT => existing confirmed-double-booking path revalidates the selected review option and persists an ACTIVE replacement fact;
- explicit REJECT => `--reject-double-booking` allows the still-unresolved session to become `TREATMENT_SUSPENDED`;
- unresolved session with no legal double-booking option may proceed directly to suspension;
- opening/closing the review UI never counts as approval or rejection.

## Native closure evidence

Native smoke date: `2026-10-26`.

### Review gate

Prepared deterministic case:

- target patient: `ΚΑΠΑΘΑΝΑΣΗ`;
- original therapist: `Βασιλειάδης`;
- original time: `08:30`;
- busy last-resort therapist: `Μηλιδάκης`;
- existing patient in that busy slot: `ΜΟΤΣΙΟΣ`.

Normal production apply stopped before suspension with:

`Double-booking review is required before treatment suspension for: planner:92:ΦΘ@2026-10-26`

This proves the suspension path cannot silently skip a valid exception review.

### ACCEPT branch

Explicit confirmation persisted an ACTIVE authoritative replacement fact:

`daily-global:planner:92:ΦΘ@2026-10-26:2026-10-26`

with:

`Βασιλειάδης -> Μηλιδάκης 08:30`

The confirmed-double-booking path remains the previously locked no-close/no-reopen flow.

### REJECT branch

Explicit rejection produced:

- persisted treatment suspensions: `1`;
- target: `planner:92:ΦΘ@2026-10-26`;
- final outcome: `Αναστολή θεραπείας`;
- unresolved sessions: `0`;
- treatment suspended sessions: `1`;
- application contract: OK with 54 components.

A suspension-only result may legitimately contain zero replacement-ledger records. In that case the production path skips the replacement-ledger installer rather than inventing a fake change record. The authoritative suspension fact and statistics remain the source of truth.

## Already verified replacement capabilities

- global scarcity-aware assignment, not sequential greedy choice;
- protection of scarce viable slots;
- placement maximization before lower-priority preferences;
- reclining/ΦΘ adjacency preference;
- original-time preference below higher-priority constraints;
- workload balancing;
- weak 13:00 avoidance;
- persistent pool/robotic service-hour policy;
- authoritative `STUDENTS` integration with distinct `STUDENT` provider kind;
- student placement/capability/daily-capacity rules;
- explicit user-confirmed double booking as last resort;
- explicit unresolved-to-`TREATMENT_SUSPENDED` fallback;
- native suspension persistence/read-back;
- statistics finalization with unresolved = 0;
- joined review/accept/reject sequencing in the production path.

## Regression guards

Relevant guards include:

- `Tests/test_apply_global_daily_replacements.py`;
- `Tests/test_apply_confirmed_double_booking.py`;
- `Tests/test_apply_daily_input_replacements_resolution.py`;
- `Tests/test_global_replacement_ledger.py`;
- `Tests/test_daily_suspension_resolution.py`.

The focused joined-flow gate passed with **21 tests** before native closure.

## Student presentation note

Active registered students are shown by real name in operational views with student-specific presentation. Numbered `φοιτ N` labels are physical placeholders only. Student identity remains authoritative through StudentID/provider kind/history.

## Closure rule

Do not reopen this replacement milestone without a concrete regression or a newly agreed business rule. New work should proceed to the next milestone while preserving this decision chain.