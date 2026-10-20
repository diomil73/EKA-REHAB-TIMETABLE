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

Restore performance is accepted for now and deferred until end-stage polish.

## Active milestone: outpatient recurring management UX

The application-style outpatient schedule manager is now implemented in code but is not yet native-verified.

Current implementation:

- `Python/rehab_excel/outpatient_schedule_manager_vba.py` owns `frmOutpatientScheduleManager` and `ShowOutpatientScheduleManager`.
- The manager shows patient, treatment, time, days and therapist.
- The worksheet row key exists only in a hidden ListBox column and is never shown or typed by the user.
- `EditOutpatientSchedule` now routes to the manager instead of asking for a worksheet row.
- The existing top-level `Διαχείριση Μητρώων` menu's outpatient button is patched to open `Προγράμματα εξωτερικών ασθενών`, so normal access remains through the application UI rather than Excel navigation.
- The manager provides `Νέο πρόγραμμα`, `Επεξεργασία`, double-click edit and `Κλείσιμο`.
- Same-row edit persistence remains protected by `target_base_entry_id`, same-patient validation and in-session update semantics.
- The combined outpatient smoke builder installs form -> edit support -> manager -> conflict guard in one Excel session.

Do not return to the old row-number InputBox flow.

## Exact resume point

1. Pull latest branch.
2. Run the focused outpatient source gate including manager, edit, in-session write, conflict and combined builder tests.
3. If green, build a fresh `OUTPATIENT_RECURRING_SMOKE.xlsm` from trusted `RERESOLVED_V90.xlsm`.
4. Require `OUTPATIENT RECURRING SMOKE BUILD OK` and `Source unchanged: True`.
5. Open only the fresh smoke copy. From the visible application UI use `Διαχείριση Μητρώων` -> `Προγράμματα εξωτερικών ασθενών`.
6. Confirm the manager shows business-facing columns only. Select an existing outpatient schedule, edit one field and save. The user must never see a worksheet row number or `outpatient:<row>` ID.
7. Reopen the manager and confirm the edited business record is present without a duplicate. Use backend/dev verification for strict same-row proof if needed; do not expose raw storage to the user.
8. On a separate fresh smoke scenario, test a known therapist/time collision and require the conflict popup/guard rather than silent double-booking.
9. Do not manually save smoke workbooks outside the tested flow.
10. After edit/conflict are native green through the manager UI, continue to the next scheduling/UI milestone. Do not return to restore performance tuning yet.

Focused gate:

```powershell
python -m pytest `
  Tests/test_outpatient_schedule_manager_vba.py `
  Tests/test_outpatient_schedule_edit_vba.py `
  Tests/test_outpatient_in_session_vba.py `
  Tests/test_outpatient_conflict_guard_vba.py `
  Tests/test_scheduling_conflict_popup_vba.py `
  Tests/test_outpatient_recurring_smoke_builder.py -q
```

Fresh smoke build:

```powershell
python Python\tools\build_outpatient_recurring_smoke.py `
  --source ".\Excel\previews\RERESOLVED_V90.xlsm" `
  --output ".\Excel\previews\OUTPATIENT_RECURRING_SMOKE.xlsm" `
  --overwrite
```

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
