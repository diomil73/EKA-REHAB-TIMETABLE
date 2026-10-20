# Project Status

Last updated: 2026-10-20

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `feature/outpatient-scheduling`.

## Current stage

Active milestone: **outpatient daily presentation cleanup, then statistics/history design checkpoint**.

CREATE is native green. The application-style manager is visible and usable in native Excel. The duplicate outpatient edit bug was reproduced and hardened. The therapist/time collision path and workload refresh were reworked and the latest fresh native smoke was reported by the user as behaving correctly. Workload visibility is now present on MASTER and the outpatient edit/conflict/workload milestone is treated as operationally green, subject to broader regression later.

The next implementation pass is presentation consistency in `THERAPIST_DAILY`, especially preventing stale special font styling from surviving outpatient edits. Immediately after that cleanup, before deep global multi-absence optimizer work, the project must hold a dedicated statistics/history design session to lock the historical facts, weekly therapist workload history, retention and dashboard semantics.

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
- collision/workload focused gate: **33 passed, 5 warnings in 4.60s**
- latest native outpatient collision + valid reassignment + workload scenario: **user reported it appears correct**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Closed: restart-safe authoritative daily state

Operational truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Replacement facts preserve immutable lifecycle history. Restore correctness, reconstruction and daily statistics parity are native green.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is only a projection/audit view.

Primary restore UX keeps the workbook open, runs Python hidden/blocking, protects rollback, imports verified state in-session and avoids visible console / close-reopen behavior.

## Outpatient recurring management + therapist workload UI

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
- `outpatient_daily_edit_vba.py` hardening that removes both plain and decorated outpatient render lines from affected old/destination daily cells before rendering once,
- workload regression coverage proving one outpatient recurring entry materializes to one session and one occupied therapist timeslot,
- `therapist_workload_strip_vba.py` with a MASTER `Φόρτος θεραπευτών` action and horizontal workload badges,
- provider-name normalization in recurring conflict checks so punctuation/accents do not create false free slots,
- automatic workload-strip refresh after outpatient changes.

### Therapist workload strip contract

The workload strip shows one badge per therapist in the form `Όνομα n/6`, where `n` is the number of occupied timeslots in the current operational day. It intentionally counts occupied timeslots, not patient lines, so a legitimate shared slot consumes one capacity unit.

Color thresholds for the normal 6-slot therapist capacity:

- **0-4 / 6: green**,
- **5 / 6: yellow/orange**,
- **6 / 6 or above: red**.

Normal users must not see raw sheet/grid/tab navigation, worksheet row numbers or `outpatient:<row>` IDs.

## Active cleanup: THERAPIST_DAILY presentation consistency

Outpatient text must not inherit stale special font styling after edits. The recurring daily style layer is being normalized so outpatient identity is communicated by its dedicated cell presentation and the patient text remains neutral black; robotic orange remains reserved for robotic treatment lines.

This is a presentation consistency gate only. It must not change scheduling truth, row identity or workload counts.

## Planned statistics/history design checkpoint

After the presentation cleanup and before deep global optimizer work, hold a dedicated design session for historical/statistical storage. At minimum it must decide:

- weekly workload history per therapist,
- daily occupied-timeslot and active-session facts,
- cancellations/no-shows/replacements/suspensions,
- historical changes of therapist assignment,
- finalized-day snapshot semantics,
- retention and comparison periods,
- dashboards and user-facing reports.

Historical metrics must derive from persisted historical facts/events, not from today's mutable projection.

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

1. Run the focused recurring-daily-style regression for neutral outpatient text and existing robotic orange semantics.
2. Verify the presentation cleanup does not regress outpatient projection/workload source tests.
3. Build one fresh native smoke only if source tests expose a presentation path needing Excel confirmation.
4. Hold the statistics/history design session, including weekly therapist workload history, before deep optimizer implementation.
5. Continue broader scheduling/UI work.
6. Implement global multi-absence optimization and persistent special-service operating hours before replacements are considered complete.
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
