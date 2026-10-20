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
- duplicate-fix focused gate: `45 passed, 4 warnings in 4.98s`
- outpatient edit after duplicate hardening: fresh native smoke appears correct

Restore performance is accepted for now and deferred until end-stage polish.

## Active milestone: therapist workload visibility

The duplicate outpatient edit path has been hardened and the latest fresh smoke appears correct. The remaining immediate operational gap is workload visibility.

Implemented in `Python/rehab_excel/therapist_workload_strip_vba.py`:

- MASTER action `Φόρτος θεραπευτών`,
- horizontal workload badges in the form `Όνομα n/6`,
- `n` counts occupied THERAPIST_DAILY timeslot cells, not patient lines,
- normal capacity threshold: green 0-4/6, yellow/orange 5/6, red 6/6 or above,
- refresh macro `ShowTherapistWorkload`,
- application-style shapes only; no raw worksheet-row concepts exposed.

The combined outpatient smoke builder now installs and verifies `modTherapistWorkloadStrip` after outpatient edit hardening.

### Why timeslots, not patient count

Operational capacity is timeslot-based. A legitimate two-patient shared slot is two sessions but one occupied timeslot, so the workload badge must show one consumed capacity unit for that clock slot.

## Exact resume point

1. Pull latest branch.
2. Run the focused workload-strip gate:

```powershell
python -m pytest `
  Tests/test_therapist_workload_strip_vba.py `
  Tests/test_outpatient_recurring_smoke_builder.py `
  Tests/test_outpatient_daily_edit_vba.py `
  Tests/test_outpatient_schedule_source.py -q
```

3. If green, build a fresh `OUTPATIENT_RECURRING_SMOKE.xlsm` from trusted `RERESOLVED_V90.xlsm`.
4. On MASTER confirm a visible `Φόρτος θεραπευτών` control exists.
5. Click it and confirm one badge per therapist appears as `Όνομα n/6`.
6. Confirm a therapist at 0-4 occupied timeslots is green, 5 is yellow/orange, 6 is red.
7. Confirm the therapist containing the outpatient edit counts that outpatient exactly once by occupied timeslot.
8. Only after this is native green, use a separate fresh smoke for therapist/time conflict verification.
9. Continue to the next unambiguous milestone without waiting for another user prompt.
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
