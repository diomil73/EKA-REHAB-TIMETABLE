# Current Handoff

Last updated: 2026-10-24

## Read order

Start with `Documentation/ONBOARDING.md`. This file is the exact resume point.

## Branch / trusted state

- Branch: `feature/outpatient-scheduling`
- Baseline `Rehab_Center_System_v27_1.xlsm` remains immutable.
- Replacement-quality milestone is CLOSED end-to-end.
- Restore correctness is closed; restore performance remains deferred.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER gate: `59 passed in 6.85s`
- latest restore profile: `16.509s`
- double-booking focused gate: `13 passed in 1.63s`
- global replacement/reclining/service-policy gate: `38 passed in 5.30s`
- student production wiring gate: `39 passed in 2.35s`
- student layout/projection gate: `19 passed in 2.35s`
- unresolved/suspension gate: `24 passed in 3.74s`
- joined replacement decision gate: `21 passed in 3.95s`
- native student replacement smoke: VERIFIED
- native suspension smoke: VERIFIED, 5 displaced / 3 replacements / 2 suspensions / unresolved 0
- native joined double-booking review gate: VERIFIED
- native explicit REJECT branch: VERIFIED, 1 persisted `Αναστολή θεραπείας`, unresolved 0, application contract OK (54 components)
- native explicit ACCEPT branch: VERIFIED, ACTIVE fact `daily-global:planner:92:ΦΘ@2026-10-26:2026-10-26`, `Βασιλειάδης -> Μηλιδάκης 08:30`

## Closed replacement decision chain

Production order is locked as:

`recurring/base -> cancellations/no-shows -> therapist absences -> surviving ACTIVE replacements -> global solve -> double-booking review -> explicit ACCEPT/REJECT -> remaining suspension -> final recompute`

Rules:

- all known therapist absences are entered before solving;
- all displaced sessions are solved together;
- scarcity/unique-slot protection outranks original-time preference;
- reclining adjacency and workload balance are optimized where feasible;
- 13:00 is only a weak soft penalty;
- pool/robotic operating hours are persistent configuration;
- therapists and students share the same solver with distinct provider kind/capability/capacity semantics;
- double-booking is last resort only and never automatic;
- if a legal double-booking option exists, production must stop for explicit review before creating suspension;
- ACCEPT persists the confirmed exception as an ACTIVE replacement fact;
- REJECT allows the orphan to become explicit `TREATMENT_SUSPENDED`;
- if no legal exception exists, suspension may occur directly;
- final operational state requires `unresolved_sessions == 0`.

Native deterministic review case on `2026-10-26`:

- patient `ΚΑΠΑΘΑΝΑΣΗ`;
- original therapist `Βασιλειάδης`;
- original time `08:30`;
- busy exception therapist `Μηλιδάκης`;
- existing patient in same slot `ΜΟΤΣΙΟΣ`.

Normal apply safety-stopped before suspension. ACCEPT and REJECT were both independently verified.

## Suspension-only ledger edge case

A final result may contain suspension facts but zero replacement-ledger records. That is valid.

The production path now skips `install_daily_change_ledger_preview(...)` when `ledger_records` is empty. It must never create a fake replacement record merely to satisfy the UI ledger installer. Suspension truth remains in `DAILY_INPUT` and daily statistics.

Relevant commits:

- `8f52c50a` production double-booking review gate
- `0695a823` sequencing regression guard
- `4659854c` workflow documentation
- `7675b4c` suspension-only ledger fix
- `c4cc54d` regression guard for empty replacement ledger
- `03bc9aa` joined workflow milestone closure documentation
- `ae211f1` project-status milestone closure

## Student presentation/lifecycle contract

- active registered students display by real name;
- `φοιτ N` labels are physical placeholders only;
- provider kind remains `STUDENT`;
- placement end removes active real-name presentation while history retains real StudentID/name;
- physical-slot overflow safety-stops.

## Therapist workload contract

MASTER badges use `Όνομα n/6`, where `n` is occupied timeslots, not patient lines:

- 0-4 green
- 5 yellow/orange
- 6+ red

History must preserve occupied timeslots and active sessions as separate metrics.

## Day-finalization/history decisions already locked

Finalization is explicit via `Κλείσιμο ημέρας`, never automatic.

Before close succeeds:

- unresolved sessions = 0;
- every concrete-date session has a final outcome;
- replacements/cancellations/no-shows/suspensions are reflected;
- final workload/statistics come from effective daily state.

History model is hybrid:

- finalized daily snapshots for reporting;
- append-only events for audit/reconstruction;
- indefinite retention;
- post-finalization correction creates a new CURRENT_FINAL version and preserves the previous one as SUPERSEDED.

Physical storage remains hidden/protected technical tables inside the same `.xlsm`.

## Patient-history requirements already locked

Preserve/derive:

- first operational entry date;
- stay duration;
- treatment/service days;
- scheduled/performed/missed sessions and structured reasons;
- expected working days excluding weekends/holidays;
- treatment counts and timeline;
- trends/totals by period;
- manual append-only 15-day evaluation milestones with `Από`, `Έως`, and green/grey/red outcome.

Patient name will later resolve the intranet/shared-folder record. Physiotherapy time opens the 15-day progress view.

## Therapist-history requirements already locked

Preserve/derive workload history, attendance/presence, dated absences and structured reasons, including the agreed Greek absence-reason list in the specialist history docs.

## Exact resume point

**Do not open another replacement vertical.** The replacement-quality milestone is closed unless a concrete regression appears.

Next milestone:

1. begin the smallest persistent statistics/history implementation slice;
2. use `Documentation/STATISTICS_HISTORY_DESIGN.md` as the design source;
3. retain explicit `Κλείσιμο ημέρας` as the only finalization action;
4. keep snapshots + append-only events in the same `.xlsm` hidden/protected storage;
5. build rule/test/Python implementation before Excel integration;
6. update docs in the same cycle;
7. keep restore-performance tuning deferred.

Likely first implementation target should be the smallest durable finalized-day snapshot/event schema and persistence gate, not charts or broad reporting UI.

## Safety reminders

- baseline immutable;
- daily facts are overlays, never destructive recurring edits;
- native Excel/COM for real `.xlsm` mutation;
- no production/staging `.xlsm` save with openpyxl;
- fresh preview/smoke copies;
- verify source hashes;
- Windows tests use `python -m pytest`;
- no manual Save outside a tested scenario;
- confirmed double-booking persistence keeps the locked no-close/no-reopen behavior.

## Documentation rule

Every agreed design/product/future-feature idea must be recorded in `Documentation/FUTURE_FEATURES.md` or the appropriate roadmap/decision document when agreed.