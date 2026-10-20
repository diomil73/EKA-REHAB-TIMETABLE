# Project Status

Last updated: 2026-10-20

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `feature/outpatient-scheduling`.

## Current stage

Active milestone: **outpatient recurring management UX**.

CREATE is native green. Same-row edit persistence is implemented, but the developer-only row-number selector must be replaced with application-style schedule management before final edit/conflict native verification.

Restore correctness is closed green. Restore performance is deferred.

## Latest confirmed checkpoints

- full regression checkpoint: **699 passed, 19 warnings**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**
- `build_master_projection`: **0.089s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session restore correctness: **native green**
- combined outpatient recurring smoke build: **native green**, trusted source unchanged
- outpatient recurring CREATE: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Closed: restart-safe authoritative daily state

Operational truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Replacement facts preserve immutable lifecycle history. Restore correctness, reconstruction and daily statistics parity are native green.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is only a projection/audit view.

Primary restore UX keeps the workbook open, runs Python hidden/blocking, protects rollback, imports verified state in-session and avoids visible console / close-reopen behavior.

## Active: outpatient recurring management UI

Current technical path already supports:

- recurring outpatient creation,
- seven suggestion modes,
- conflict backend and popup,
- stable `target_base_entry_id`,
- same-row edit instead of duplicate append,
- same-patient target validation,
- current-date daily-view cleanup/re-render when edited.

Open UX gate:

- the existing `EditOutpatientSchedule` row-number InputBox is developer scaffolding only,
- normal users do not see raw sheets/grid/tabs,
- `OUTPATIENT_SCHEDULE` remains hidden/protected storage,
- build a business-facing management/selection UI with patient, treatment, time, days and therapist,
- selection must carry the internal target invisibly.

Do not ask the user to continue edit testing through worksheet row numbers.

## Open: replacement quality under multiple therapist absences

Replacement lifecycle correctness is green, but global assignment quality remains open.

Required global behavior:

- enter all known therapist absences first,
- solve all displaced patients together,
- maximize placements / minimize unresolved orphans,
- protect unique viable slots from flexible patients,
- prefer original time when higher-priority constraints permit,
- preserve ΦΘ/reclining proximity where feasible,
- avoid significant provider-load imbalance,
- respect patient absences and other treatments,
- add persistent pool/robotic operating-hour configuration,
- treat 13:00 as a weak soft preference to avoid,
- permit double-booking/over-capacity only through explicit last-resort user approval.

Canonical scarcity example: A can use 09:00/10:00/11:00, B only 09:00, so protect 09:00 for B when a complete feasible solution exists.

This must be addressed before replacements are considered complete.

## Deferred: restore performance

Current end-user restore duration in the low-to-mid 20-second range is accepted for this stage. Do not resume performance work unless there is a regression or final-polish phase.

## Immediate roadmap

1. Build outpatient schedule management/selection UI.
2. Add pure mapping/selection tests that keep internal row IDs hidden from users.
3. Build a fresh smoke copy and native-test create -> reopen -> edit same authoritative entry with no duplicate.
4. Native-test therapist/time conflict on a separate fresh scenario.
5. Continue broader scheduling/UI work.
6. Before closing replacements, implement global multi-absence optimization and persistent special-service operating hours.
7. Revisit restore performance near completion.

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
