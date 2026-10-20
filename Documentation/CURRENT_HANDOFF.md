# Current Handoff

Last updated: 2026-10-20

## Read order

For any new developer, new ChatGPT context, or long-gap return, read `Documentation/ONBOARDING.md` first. This file is only the exact current resume point.

## Branch / trusted source

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER projection focused gate: `59 passed in 6.85s`
- latest direct restore backend profile: `16.509s total`
- latest `build_master_projection`: `0.089s`
- replacement lifecycle V87-V92: native green
- per-row / selected / restore-all in-session restore correctness: native green
- combined outpatient recurring smoke build: native green, trusted source unchanged
- outpatient recurring CREATE: native green
- outpatient schedule manager: native visible and usable
- outpatient edit: not green; native test exposed duplicate `THERAPIST_DAILY` placement after edit

Restore performance is accepted for now and deferred until end-stage polish.

## Active milestone: outpatient recurring edit correctness

The application-style outpatient schedule manager is implemented and native-visible.

Current implementation:

- `Python/rehab_excel/outpatient_schedule_manager_vba.py` owns `frmOutpatientScheduleManager` and `ShowOutpatientScheduleManager`.
- The manager shows patient, treatment, time, days and therapist.
- The worksheet row key exists only in a hidden ListBox column and is never shown or typed by the user.
- `EditOutpatientSchedule` routes to the manager instead of asking for a worksheet row.
- The existing top-level `Διαχείριση Μητρώων` menu's outpatient button opens `Προγράμματα εξωτερικών ασθενών`.
- Same-row authoritative edit semantics remain protected by `target_base_entry_id` and same-patient validation.

### Blocking native bug

A native edit that changes an outpatient placement can leave the patient visible in the old `THERAPIST_DAILY` slot and also add the patient to the new slot. The result is a duplicated patient in the daily therapist view.

Treat this as a correctness blocker:

- edit is not native green,
- do not trust the affected `THERAPIST_DAILY` view as proof of placement correctness,
- do not trust therapist workload/capacity numbers derived from a duplicated daily projection,
- do not continue the conflict native gate on the affected smoke.

Likely weakness: the current in-session edit path tries to remove the old outpatient line by matching visible text, then appends the new placement. This is too fragile. The preferred direction is deterministic reconstruction of the affected daily slot(s) from authoritative daily truth so one base entry produces exactly one effective daily placement.

Do not return to the old row-number InputBox flow.

## Exact resume point

1. Pull latest branch.
2. Add a regression that reproduces the duplicate outpatient after edit.
3. Refactor the affected-slot refresh so it is reconstructed deterministically from authoritative daily truth rather than relying on text-line deletion followed by append.
4. Add a regression proving exactly one outpatient placement remains after edit: old slot absent, new slot present once.
5. Add a workload/capacity regression proving the outpatient counts exactly once for the therapist after edit.
6. Run the focused outpatient edit/source/in-session tests.
7. Build a fresh `OUTPATIENT_RECURRING_SMOKE.xlsm` from trusted `RERESOLVED_V90.xlsm`.
8. Native-test through `Διαχείριση Μητρώων` -> `Προγράμματα εξωτερικών ασθενών`: edit one field, save, reopen, confirm one record and one daily placement only.
9. Only after edit + workload are native green, use a separate fresh smoke for therapist/time conflict verification.
10. Do not manually save smoke workbooks outside the tested flow.
11. Do not return to restore performance tuning yet.

## Open replacement-quality milestone

The replacement lifecycle is correct, but global assignment quality remains open.

All known therapist absences for a date must be entered first. Then all displaced sessions are solved together, not greedily.

Locked highlights:

- maximize placed displaced patients and minimize unresolved orphans,
- protect unique viable slots from flexible patients,
- if A can use 09:00/10:00/11:00 and B only 09:00, protect 09:00 for B,
- prefer original time when it does not damage a higher-priority constraint,
- keep ΦΘ and reclining treatment close where feasible,
- keep workload balanced and avoid a two-timeslot imbalance where feasible,
- 13:00 is only a weak soft preference to avoid,
- pool/robotic operating hours must become persistent configuration,
- double-booking/over-capacity is last resort only and requires explicit user approval.

## Locked daily truth

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Recomputation order:

1. recurring/base schedule,
2. cancellations/no-shows release capacity,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply surviving ACTIVE replacements,
6. resolve remaining orphans through placement or treatment suspension,
7. only then expose capacity to new patients.

Every orphan on the concrete date resolves to a valid therapist/time placement or `Αναστολή θεραπείας`. Finalized day requires unresolved=0.

## Safety reminders

- baseline is immutable,
- fresh preview/smoke copies only,
- preserve VBA and source hash,
- do not save production/staging `.xlsm` with openpyxl,
- Windows tests use `python -m pytest`, never bare `pytest`,
- no manual Save during independent smoke scenarios.

## Documentation rule

`Documentation/ONBOARDING.md` is now the single front door. Update the specialist source of truth for each meaningful change and update ONBOARDING too when the project-wide mental model changes.
