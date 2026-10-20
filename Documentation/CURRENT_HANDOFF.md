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
- outpatient edit: not green yet; duplicate-fix code is implemented but awaits focused + native verification

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

### Duplicate daily projection bug and implemented hardening

Native edit showed the same outpatient twice in `THERAPIST_DAILY`. The identified mechanism is that normal daily rendering may decorate the patient line, for example `Όνομα [Καθ/να]`, while the old edit cleanup removed only a line exactly equal to `Όνομα`. The old rendered line therefore survived and the in-session edit appended a second plain line.

Implemented in `Python/rehab_excel/outpatient_daily_edit_vba.py`:

- patch the installed `modOutpatientInSession` after the normal conflict/in-session installer,
- on edit, clean the old affected daily cell before refresh,
- if therapist/time changes, clean the destination cell too,
- recognize both plain patient lines and known decorated patient-line forms such as `Όνομα [..]`, `Όνομα (..)` and replacement-arrow suffixes,
- render only once after cleanup,
- keep this hardening isolated from authoritative storage semantics.

The combined outpatient smoke builder now installs this hardening after the conflict guard and before verification/save.

This remains a transitional in-session hardening layer. Long-term rendering direction is deterministic reconstruction from authoritative daily truth rather than text-oriented cell surgery.

### Workload verification

A new Python regression materializes one outpatient recurring entry on an applicable date and verifies `calculate_therapist_workload` returns exactly:

- `active_sessions == 1`
- `active_timeslots == 1`

So the scheduling model has an explicit one-outpatient/one-load contract. Native Excel must now prove its projection matches that model.

Do not return to the old row-number InputBox flow.

## Exact resume point

1. Pull latest branch.
2. Run the focused duplicate-fix gate:

```powershell
python -m pytest `
  Tests/test_outpatient_daily_edit_vba.py `
  Tests/test_outpatient_schedule_source.py `
  Tests/test_outpatient_schedule_registration.py `
  Tests/test_outpatient_conflict_guard_vba.py `
  Tests/test_outpatient_schedule_edit_vba.py `
  Tests/test_outpatient_schedule_manager_vba.py `
  Tests/test_outpatient_recurring_smoke_builder.py -q
```

3. If green, build a fresh `OUTPATIENT_RECURRING_SMOKE.xlsm` from trusted `RERESOLVED_V90.xlsm`.
4. Native-test through `Διαχείριση Μητρώων` -> `Προγράμματα εξωτερικών ασθενών`: edit only the time of the existing outpatient.
5. Confirm the old daily placement is gone and the destination contains the outpatient exactly once.
6. Reopen the manager and confirm the schedule entry itself is still single and shows the new time.
7. Confirm therapist load semantics remain one outpatient session / one occupied timeslot. Do not infer correctness merely from duplicate-looking text.
8. If this is green, clean remaining outpatient daily styling issues.
9. Then use a separate fresh smoke for therapist/time conflict verification.
10. Continue to the next unambiguous milestone without waiting for an extra user prompt.
11. Do not manually save smoke workbooks outside the tested flow.
12. Do not return to restore performance tuning yet.

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
