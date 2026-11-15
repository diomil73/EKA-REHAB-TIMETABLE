# Project Status

Last updated: 2026-10-24

## Entry point

Start from `Documentation/ONBOARDING.md`. Use `Documentation/CURRENT_HANDOFF.md` for the exact resume point.

Branch: `feature/outpatient-scheduling`.

## Current stage

**Replacement-quality milestone: CLOSED.**

The joined production path is now native-verified end-to-end:

`global solve -> double-booking review -> explicit ACCEPT or REJECT -> remaining suspension -> final recompute`

The next implementation milestone is the smallest persistent statistics/history slice built on the already locked day-finalization design.

Restore correctness remains closed green. Restore performance remains deferred to final polish.

## Latest confirmed checkpoints

- full regression checkpoint: **699 passed, 19 warnings**
- OOXML MASTER projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**
- replacement lifecycle V87-V92: **native green**
- per-row / selected / restore-all authoritative restore: **native green**
- outpatient recurring CREATE/edit/manager: **native green / usable**
- THERAPIST_DAILY cleanup gate: **27 passed in 3.09s**
- double-booking focused gate: **13 passed in 1.63s**, native V5 green
- global replacement / reclining / service-policy gate: **38 passed in 5.30s**
- student production wiring gate: **39 passed in 2.35s**
- student layout/projection gate: **19 passed in 2.35s**
- native student replacement smoke: **VERIFIED**
- unresolved-to-suspension gate: **24 passed in 3.74s**
- native suspension-resolution smoke: **VERIFIED**, 5 displaced, 3 replacements, 2 suspensions, unresolved 0
- joined double-booking sequencing gate: **21 passed in 3.95s**
- native joined review gate: **VERIFIED** on `2026-10-26`
- native explicit REJECT branch: **VERIFIED**, 1 persisted suspension, unresolved 0, app contract OK (54 components)
- native explicit ACCEPT branch: **VERIFIED**, ACTIVE fact for `planner:92:ΦΘ@2026-10-26`, `Βασιλειάδης -> Μηλιδάκης 08:30`

Python 3.14 `ZipFile.__del__` warnings remain non-blocking technical debt.

## Closed: replacement-quality milestone

The production replacement path now satisfies the locked contract:

- all displaced sessions are solved together, not greedily;
- scarcity and unique-slot protection are respected;
- original time is preferred only below higher-priority feasibility/scarcity constraints;
- ΦΘ/reclining adjacency is optimized where feasible;
- workload spread is considered;
- 13:00 is only a weak soft penalty;
- pool/robotic service windows are persistent configuration;
- therapists and students share one solver while preserving distinct provider kind/capability/capacity rules;
- double-booking is last resort only and requires explicit user approval;
- valid review options block automatic suspension until the operator decides;
- explicit rejection allows suspension fallback;
- every remaining orphan becomes explicit `TREATMENT_SUSPENDED` before finalization;
- finalized effective daily state requires unresolved = 0.

Native deterministic case used patient `ΚΑΠΑΘΑΝΑΣΗ`, original therapist `Βασιλειάδης`, slot `08:30`, with busy therapist `Μηλιδάκης` already treating `ΜΟΤΣΙΟΣ`.

The normal path safety-stopped for review. ACCEPT persisted the approved double booking as an ACTIVE replacement fact. REJECT persisted `Αναστολή θεραπείας` and finalized with unresolved 0.

A suspension-only final outcome may legitimately have no replacement ledger records. The production path now skips the replacement-ledger installer in that case instead of creating a fake record.

## Closed: restart-safe authoritative daily state

Operational truth remains:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

The base schedule is never destructively rewritten by daily absences, replacements or suspensions.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a projection/audit surface only.

Primary restore UX keeps the workbook open and preserves the locked no-close/no-reopen behavior for confirmed double-booking persistence.

## Student provider presentation/lifecycle contract

Active students display by real name in operational views with student-specific presentation. `φοιτ N` labels are physical placeholders only. Historical identity remains the real StudentID/name.

## Statistics/history design checkpoint

The architecture is already locked for:

- explicit `Κλείσιμο ημέρας`;
- finalized daily snapshots + append-only events;
- indefinite retention;
- versioned amendments with SUPERSEDED/CURRENT_FINAL chain;
- therapist occupied-timeslot and active-session history as separate measures;
- patient history/timeline;
- therapist attendance/absence history;
- filtered tables plus charts driven by the same finalized truth.

## Immediate roadmap

1. Start the smallest persistent statistics/history implementation slice.
2. Preserve the replacement decision chain as a closed dependency, reopening only on regression.
3. Keep restore performance deferred until final polish.
4. Keep application-style UI rules intact: normal users use forms/buttons, not raw technical sheets/IDs.

## Locked daily rules

1. recurring/base schedule;
2. cancellations/no-shows release capacity;
3. therapist absences;
4. identify displaced/orphan sessions;
5. apply surviving ACTIVE replacements;
6. global solve remaining displaced sessions;
7. review any legal last-resort double-booking option;
8. explicit ACCEPT or REJECT;
9. remaining unresolved -> `Αναστολή θεραπείας`;
10. final recompute requires unresolved = 0;
11. only then expose capacity to new patients.

Capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Safety constraints

- baseline `.xlsm` is immutable;
- preview/smoke workflows use fresh copies;
- preserve VBA and source hash;
- keep write zones narrow;
- no production/staging `.xlsm` save with openpyxl;
- no manual Save in independent smoke scenarios;
- Windows tests use `python -m pytest`.

## Documentation discipline

Meaningful changes must update the relevant specialist document in the same cycle. Agreed future design/product ideas belong in `Documentation/FUTURE_FEATURES.md` or the appropriate roadmap/decision document.