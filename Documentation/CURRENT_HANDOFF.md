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
- THERAPIST_DAILY presentation cleanup focused gate: `27 passed in 3.09s`

Restore performance is accepted for now and deferred until end-stage polish.

## Current milestone: statistics/history design checkpoint

The outpatient edit/conflict/workload path is operationally green enough to move forward. The THERAPIST_DAILY presentation cleanup source gate is also green.

The next work is not another ad hoc statistic. Before implementing persistence, explicitly design the historical fact model so later schedule edits cannot rewrite yesterday's history.

The user explicitly wants weekly workload history per therapist plus broader statistics/history.

## Therapist workload contract already locked

MASTER exposes `Φόρτος θεραπευτών` with one badge per therapist in the form `Όνομα n/6`.

`n` is occupied timeslots, not patient-line count. Color thresholds:

- 0-4/6 green,
- 5/6 yellow/orange,
- 6/6 or above red.

For history, preserve both concepts separately:

- occupied timeslots, for capacity utilization,
- active sessions, for actual patient/session workload.

A legitimate shared slot can therefore be 1 occupied timeslot but 2 active sessions.

## Statistics/history design topics to lock now

At minimum decide:

- daily historical row/fact grain,
- weekly therapist workload history,
- occupied-timeslot vs active-session metrics,
- planned vs actually performed workload,
- cancellations/no-shows/patient absences,
- replacements and replacement-provider load,
- treatment suspensions,
- therapist absence days/hours,
- permanent therapist-assignment changes,
- finalized-day snapshot semantics,
- whether corrections after finalization are append-only revisions rather than destructive edits,
- retention horizon and comparison windows,
- user-facing dashboard/report layout,
- privacy boundary for patient-level historical reporting.

Historical metrics must come from persisted historical facts/events, never from today's mutable THERAPIST_DAILY or MASTER projection.

## Exact resume point

1. Treat the `27 passed in 3.09s` presentation gate as green.
2. Hold the statistics/history design discussion now.
3. Lock the smallest durable historical schema before implementing Excel UI for reports.
4. Implement weekly therapist workload history as the first concrete slice once the contract is agreed.
5. Then continue broader scheduling/UI work and later the global multi-absence optimizer.
6. Do not manually save smoke workbooks outside a tested flow.
7. Do not return to restore performance tuning yet.

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
