# Project Status

Last updated: 2026-10-20

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `feature/outpatient-scheduling`.

## Current stage

Active milestone: **therapist workload visibility + remaining outpatient native gates**.

CREATE is native green. The application-style manager is visible and usable in native Excel. The duplicate outpatient edit bug was reproduced, hardened, covered by focused regression, and a fresh native smoke now appears correct: the edited outpatient no longer visibly duplicates. The next operational gap is that therapist workload is not visible to the user, so a dedicated MASTER workload strip is now implemented in code and awaits focused/native verification.

Restore correctness is closed green. Restore performance is deferred.

## Latest confirmed checkpoints

- full regression checkpoint: **699 passed, 19 warnings**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**
- `build_master_projection`: **0.089s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session restore correctness: **native green**
- previous combined outpatient recurring smoke build: **native green**, trusted source unchanged
- outpatient recurring CREATE: **native green**
- outpatient manager: **native visible and usable**
- duplicate-fix focused gate: **45 passed, 4 warnings in 4.98s**
- outpatient EDIT after duplicate hardening: **fresh native smoke appears correct; formal workload visibility check follows**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Closed: restart-safe authoritative daily state

Operational truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Replacement facts preserve immutable lifecycle history. Restore correctness, reconstruction and daily statistics parity are native green.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is only a projection/audit view.

Primary restore UX keeps the workbook open, runs Python hidden/blocking, protects rollback, imports verified state in-session and avoids visible console / close-reopen behavior.

## Active: outpatient recurring management + therapist workload UI

Current technical path supports:

- recurring outpatient creation,
- seven suggestion modes,
- conflict backend and popup,
- stable `target_base_entry_id`,
- same-row edit semantics in authoritative storage,
- same-patient target validation,
- application-style `frmOutpatientScheduleManager` listing patient, treatment, time, days and therapist,
- hidden internal row mapping inside the manager rather than user-entered row numbers,
- `EditOutpatientSchedule` routed through the manager,
- the existing `Διαχείριση Μητρώων` outpatient action routed to `Προγράμματα εξωτερικών ασθενών`,
- create/edit/close and double-click edit from the manager,
- combined smoke installation of form, edit support, manager and conflict guard in one Excel session,
- `outpatient_daily_edit_vba.py` hardening that removes both plain and decorated outpatient render lines such as `Όνομα [Καθ/να]` from affected old/destination daily cells before rendering once,
- workload regression coverage proving one outpatient recurring entry materializes to one session and one occupied therapist timeslot,
- new `therapist_workload_strip_vba.py` with a MASTER `Φόρτος θεραπευτών` action and horizontal workload badges.

### Therapist workload strip contract

The workload strip shows one badge per therapist in the form `Όνομα n/6`, where `n` is the number of occupied timeslot cells in the current `THERAPIST_DAILY` projection. It intentionally counts occupied timeslots, not patient lines, so a legitimate shared slot consumes one capacity unit.

Color thresholds are now locked for the normal 6-slot therapist capacity:

- **0-4 / 6: green**,
- **5 / 6: yellow/orange**,
- **6 / 6 or above: red**.

The action lives on the MASTER as an application-style control, not as raw worksheet navigation. The combined outpatient smoke builder installs this workload strip so it can be verified in the same fresh native scenario.

Normal users must not see raw sheet/grid/tab navigation, worksheet row numbers or `outpatient:<row>` IDs.

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

1. Run the focused workload-strip + smoke-builder source gate.
2. Build a fresh combined outpatient smoke from trusted `RERESOLVED_V90.xlsm`.
3. On MASTER, verify the visible `Φόρτος θεραπευτών` action and the horizontal workload badges.
4. Verify a therapist with the outpatient edit still counts the outpatient exactly once by occupied timeslot, and that the badge refresh reflects the current `THERAPIST_DAILY` state.
5. Verify color thresholds: green through 4/6, yellow/orange at 5/6, red at 6/6.
6. Only after workload UI is native green, use a separate fresh smoke for therapist/time conflict verification.
7. Continue broader scheduling/UI work without waiting for another prompt when the next step is unambiguous.
8. Before closing replacements, implement global multi-absence optimization and persistent special-service operating hours.
9. Revisit restore performance near completion.

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
