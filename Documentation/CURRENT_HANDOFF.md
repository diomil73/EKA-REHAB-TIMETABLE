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
- collision/workload focused gate: `33 passed, 5 warnings in 4.60s`
- latest fresh native outpatient collision + valid reassignment + workload scenario: user reported it appears correct

Restore performance is accepted for now and deferred until end-stage polish.

## Current milestone: THERAPIST_DAILY presentation cleanup

The outpatient edit/conflict/workload path is operationally green enough to move forward. The immediate cleanup is visual consistency in `THERAPIST_DAILY` without changing scheduling truth.

`Python/rehab_excel/recurring_daily_style.py` is now being normalized so:

- ordinary recurring text is black,
- outpatient recurring text is also neutral black,
- outpatient identity comes from the dedicated outpatient cell presentation layer rather than a special font color,
- robotic treatment lines retain orange precedence,
- stale inherited special font color after an outpatient edit is not allowed.

This cleanup follows the same rule as the in-session outpatient renderer, which already writes non-robotic outpatient text in black.

## Therapist workload contract

MASTER exposes `Φόρτος θεραπευτών` with one badge per therapist in the form `Όνομα n/6`.

`n` is occupied timeslots, not patient-line count. Color thresholds:

- 0-4/6 green,
- 5/6 yellow/orange,
- 6/6 or above red.

The strip refreshes after outpatient changes and provider names are normalized for collision checks so accents/punctuation do not create false free slots.

## Statistics/history checkpoint

The user wants weekly workload history per therapist plus broader statistics/history. Do not design or persist this ad hoc while scheduling semantics are still moving.

The correct design checkpoint is immediately after the current presentation cleanup and before deep global multi-absence optimizer implementation. At that point explicitly discuss and lock:

- weekly therapist workload history,
- daily occupied-timeslot and active-session facts,
- cancellations/no-shows/replacements/suspensions,
- therapist-assignment change history,
- finalized-day snapshots,
- retention/comparison windows,
- user-facing dashboards/reports.

Historical metrics must come from persisted historical facts/events, not from today's mutable projection.

## Exact resume point

1. Pull latest branch.
2. Run the focused presentation gate:

```powershell
python -m pytest `
  Tests/test_recurring_daily_style.py `
  Tests/test_outpatient_presentation.py `
  Tests/test_outpatient_daily_edit_vba.py `
  Tests/test_therapist_workload_strip_vba.py -q
```

3. If green, continue to the statistics/history design checkpoint without waiting for another prompt.
4. Do not manually save smoke workbooks outside a tested flow.
5. Do not return to restore performance tuning yet.

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
