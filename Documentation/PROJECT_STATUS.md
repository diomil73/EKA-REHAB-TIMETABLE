# Project Status

Last updated: 2026-10-20

## Entry point

New developers and new contexts start from `Documentation/ONBOARDING.md`. This file tracks milestone state, risks and immediate roadmap only.

Branch: `feature/outpatient-scheduling`.

## Current stage

Active milestone: **outpatient recurring management UX native verification**.

CREATE is native green. The application-style manager is now visible and usable in native Excel. Native edit testing exposed a blocking correctness issue: an edited outpatient can appear twice in `THERAPIST_DAILY`, so the edit gate is not green yet and therapist workload derived from that view must not be trusted until the duplication path is fixed and verified.

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
- outpatient EDIT: **blocked by duplicate daily projection**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## Closed: restart-safe authoritative daily state

Operational truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Replacement facts preserve immutable lifecycle history. Restore correctness, reconstruction and daily statistics parity are native green.

## Closed: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is only a projection/audit view.

Primary restore UX keeps the workbook open, runs Python hidden/blocking, protects rollback, imports verified state in-session and avoids visible console / close-reopen behavior.

## Active: outpatient recurring management UI

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
- combined smoke installation of form, edit support, manager and conflict guard in one Excel session.

### Blocking native correctness issue

Native edit verification revealed that the outpatient patient may remain in the old `THERAPIST_DAILY` placement and also be rendered in the new placement, producing a duplicate visible patient.

This is a correctness issue, not cosmetic. Until it is fixed:

- outpatient edit is **not native green**,
- `THERAPIST_DAILY` may not be used as proof of correct outpatient placement after edit,
- therapist workload/capacity derived from that duplicated daily projection must be treated as unverified,
- no further conflict or workload native acceptance should be recorded on top of the affected smoke.

The fix should prefer deterministic reconstruction of the affected daily placement from authoritative daily truth instead of fragile text-line removal followed by append.

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

1. Reproduce the outpatient edit duplicate deterministically in tests.
2. Replace fragile old-slot text removal/append behavior with deterministic affected-slot reconstruction from authoritative daily truth.
3. Add regression coverage proving one outpatient base entry produces exactly one visible daily placement after edit.
4. Add a workload/capacity regression proving the outpatient contributes exactly once to the therapist's operational load after edit.
5. Build a fresh combined outpatient smoke and native-test edit again: old placement gone, new placement present once, no duplicate, correct styling.
6. Only after that gate is green, native-test therapist/time conflict on a separate fresh scenario.
7. Continue broader scheduling/UI work.
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
