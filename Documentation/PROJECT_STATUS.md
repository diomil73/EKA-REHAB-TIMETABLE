# Project Status

Last updated: 2026-10-24

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `feature/outpatient-scheduling`.

## Current stage

Active milestone: **replacement-quality closure pass**.

Restore correctness is closed green. Restore performance is deferred.

Outpatient recurring CREATE/edit/manager, therapist workload refresh, THERAPIST_DAILY presentation cleanup, confirmed double-booking, global replacement solving, student-provider wiring and explicit suspension fallback are all operationally verified on their focused/native paths.

The current task is no longer to add obvious missing replacement primitives. It is to reassess the locked replacement-quality checklist against the now-implemented system and identify any remaining genuine gap before declaring the replacement milestone closed.

## Latest confirmed checkpoints

- full regression checkpoint: **699 passed, 19 warnings**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**
- `build_master_projection`: **0.089s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session restore correctness: **native green**
- combined outpatient recurring smoke build: **native green**, trusted source unchanged
- outpatient recurring CREATE: **native green**
- outpatient manager: **native visible and usable**
- duplicate-fix focused gate: **45 passed, 4 warnings in 4.98s**
- collision/workload focused gate: **33 passed, 5 warnings in 4.60s**
- THERAPIST_DAILY presentation cleanup gate: **27 passed in 3.09s**
- double-booking focused gate: **13 passed in 1.63s**, native V5 green
- global replacement / reclining / service-policy gate: **38 passed in 5.30s**
- student production wiring gate: **39 passed in 2.35s**
- student layout/projection gate: **19 passed in 2.35s**
- native student replacement smoke: **VERIFIED**, 2 ACTIVE student replacement facts, THERAPIST_DAILY OK
- unresolved-to-suspension focused gate: **24 passed in 3.74s**
- native suspension-resolution smoke: **VERIFIED** on `2026-10-26`; 5 displaced sessions, 3 replacement assignments, 2 persisted treatment suspensions, unresolved 0, input unchanged, application contract OK (54 components)

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Closed: restart-safe authoritative daily state

Operational truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Replacement facts preserve immutable lifecycle history. Restore correctness, reconstruction and daily statistics parity are native green.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is only a projection/audit view.

Primary restore UX keeps the workbook open, runs Python hidden/blocking, protects rollback, imports verified state in-session and avoids visible console / close-reopen behavior.

## Closed: unresolved orphan finalization

Global replacement solver leftovers no longer block the day solely because they cannot be placed.

The finalization path is now:

1. solve all displaced sessions globally,
2. persist valid replacement assignments,
3. convert remaining unplaceable sessions to explicit `TREATMENT_SUSPENDED` daily cancellations,
4. recompute effective daily state,
5. require unresolved sessions = 0,
6. persist suspension facts into `DAILY_INPUT` and verify read-back,
7. allow daily ledger completion only when every solver leftover is covered by an explicit suspension fact.

Native smoke verified this with 3 replacements + 2 suspensions and zero unresolved sessions.

## Outpatient recurring management + therapist workload UI

Current technical path supports recurring outpatient creation/edit, application-style management, conflict popup/guard, stable same-row identity, daily projection cleanup, workload badges and automatic refresh after outpatient changes.

### Therapist workload strip contract

The workload strip shows one badge per therapist in the form `Όνομα n/6`, where `n` is the number of occupied timeslots in the current operational day. It intentionally counts occupied timeslots, not patient lines, so a legitimate shared slot consumes one capacity unit.

Color thresholds for the normal 6-slot therapist capacity:

- **0-4 / 6: green**,
- **5 / 6: yellow/orange**,
- **6 / 6 or above: red**.

Normal users must not see raw sheet/grid/tab navigation, worksheet row numbers or `outpatient:<row>` IDs.

## Student provider presentation/lifecycle contract

Active students display by their real names in THERAPIST_DAILY and workload surfaces with student-specific visual treatment. Numbered `φοιτ N` headers are physical placeholders only. When placement ends, the real name disappears from active operational views and the physical slot returns to its neutral placeholder. Historical identity remains the real StudentID/name.

## Replacement-quality checklist under reassessment

Implemented/verified areas already include:

- global scarcity-aware solving across all displaced sessions,
- unique-slot protection / placement maximization,
- original-time preference below higher-priority constraints,
- reclining adjacency preference,
- workload-balance preference,
- 13:00 weak avoidance,
- persistent pool/robotic service-hours policy,
- authoritative student-provider path,
- explicit suspension fallback for unplaceable sessions,
- explicit user approval for confirmed double-booking.

Remaining work should be chosen only after comparing this implementation against the complete locked contract. Do not reopen already-green verticals without a concrete regression or missing rule.

## Statistics/history design checkpoint

The architecture is locked for explicit `Κλείσιμο ημέρας`, hybrid finalized snapshots + append-only events, indefinite retention, versioned amendments, therapist occupied-timeslot vs active-session history, patient timeline/history, and user-facing reports.

Implementation remains deferred until the replacement-quality milestone is explicitly closed.

## Immediate roadmap

1. Reassess every remaining replacement-quality requirement against current code/tests/native smokes.
2. Implement only the smallest genuinely missing replacement rule, if any.
3. Run the relevant focused regression/native smoke and update documentation.
4. When the replacement checklist is fully satisfied, explicitly close the replacement milestone.
5. Resume the smallest persistent statistics/history implementation slice after replacement closure.
6. Keep restore performance deferred until final polish unless a regression appears.

## Locked daily rules

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. unresolved orphans resolve through valid placement or treatment suspension,
7. only then expose capacity to new patients.

Capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Safety constraints

- baseline `.xlsm` is immutable,
- preview/smoke workflows use fresh copies,
- preserve VBA and source hash,
- keep write zones narrow,
- bulk restore validates the entire requested set before mutation,
- no production/staging `.xlsm` save with openpyxl,
- no manual Save in independent smoke scenarios,
- Windows tests use `python -m pytest`.

## Documentation discipline

`Documentation/ONBOARDING.md` is the single front door. Specialist files keep their own source-of-truth roles. Meaningful changes must update the relevant specialist document in the same cycle, and ONBOARDING when the project-wide mental model changes.
