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

Same-row edit persistence is already hardened in code:

- `target_base_entry_id` is preserved through backend conflict validation and final in-session write,
- edit updates the same `outpatient:<row>` instead of appending,
- same-patient validation is enforced,
- old current-date `THERAPIST_DAILY` placement is removed before the edited placement is rendered when needed.

However, the existing `EditOutpatientSchedule` InputBox that asks for an `OUTPATIENT_SCHEDULE` worksheet row is developer/test scaffolding only and is not acceptable as final UI.

Locked UI direction:

- `OUTPATIENT_SCHEDULE` stays hidden/protected authoritative storage,
- normal users must not know sheet names, row numbers or `outpatient:<row>` IDs,
- create/edit must be available through an application-style management flow,
- the management UI should list patient, treatment, time, days and therapist,
- the selected business-facing row carries the internal target invisibly,
- do not ask the user to continue native edit testing through the row-number InputBox.

## Exact resume point

1. Pull latest branch.
2. Implement an application-style outpatient schedule management/selection UI.
3. Add pure tests proving the selected visible entry maps to the intended `target_base_entry_id` without exposing worksheet row numbers.
4. Build a fresh smoke copy.
5. Native-test: create an outpatient schedule, reopen it through the management UI, edit one field, save, and confirm same authoritative entry with no duplicate.
6. On a separate fresh scenario, test known therapist/time collision and require the conflict popup/guard rather than silent double-booking.
7. Do not manually save smoke workbooks outside the tested flow.
8. After create/edit/conflict are native green through the application-style UI, continue to the next scheduling/UI milestone.
9. Before replacements are considered complete, implement the global multi-absence optimizer described in `Documentation/DECISIONS.md` and `Documentation/ONBOARDING.md`.
10. Do not return to restore performance tuning unless there is a regression or final-polish phase.

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
