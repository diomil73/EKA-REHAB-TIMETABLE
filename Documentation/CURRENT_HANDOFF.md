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

The active work is now the historical fact model. Historical statistics must not be reconstructed later from today's mutable schedule projection.

The user explicitly wants weekly workload history per therapist plus broader statistics/history and patient-level analysis.

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

Primary historical emphasis is the actual final operational workload. Planned workload may also be retained as secondary comparison context.

## Day-finalization decision

Historical daily facts become final only through an explicit application action: **`Κλείσιμο ημέρας`**.

Do not auto-finalize based on clock time or workbook close.

Before the action succeeds:

- unresolved/orphan sessions must be zero,
- every concrete-date session must have a final outcome,
- replacements/cancellations/no-shows/suspensions must already be reflected,
- final workload/statistics must come from the effective daily state.

Successful close persists the finalized daily snapshot used by weekly history.

## Persistence model decision

History is **hybrid**:

- finalized daily snapshots are the fast reporting source for weekly/monthly history,
- append-only historical events preserve the audit trail and support later corrections/reconstruction.

The snapshot is not allowed to replace or erase event history.

At minimum the event model must be able to represent absences, cancellations/no-shows, replacements, suspensions, approved over-capacity/double-booking exceptions, relevant permanent therapist-assignment changes, day finalization and post-finalization amendments.

## Physical storage decision

Historical snapshots, therapist/day workload rows and append-only events stay inside the same `.xlsm` in hidden/protected technical tables.

Normal operation must not depend on external history files. Users access history through application forms/reports, not raw technical sheets. Patient-level history must not be automatically exported to the repository or uncontrolled external artifacts.

## Finalized-day amendment decision

Closed days are corrected through an explicit application action such as **`Διόρθωση κλεισμένης ημέρας`**, not through destructive reopen-and-rewrite.

A valid correction:

- appends an amendment event,
- applies corrected facts through the validated daily business rules,
- must pass the finalization gate again,
- creates a new finalized snapshot version,
- preserves the previous version as `SUPERSEDED`,
- links the new snapshot to the previous one,
- keeps reports on the latest `CURRENT_FINAL` version,
- keeps the full version chain available for audit.

A failed correction must not partially replace the current-final historical state.

## Retention decision

Historical snapshots and append-only events are retained indefinitely by default for the life of the system. There is no automatic aging/deletion policy. Any future archive operation must be explicit, controlled and auditable.

## Patient-level analysis decision

Patient-level historical analysis is required.

For a selected patient, the application must be able to derive/display at least scheduled vs performed sessions, absences/no-shows, cancellations/postponements, treatment suspensions, replacements with effective therapist/time, therapist-assignment changes and period trends/totals.

Stable `PatientID` remains the internal key. The UI shows normal patient identity/business labels. Patient history must be based on finalized snapshots/events, not on the patient's current recurring schedule.

## Dashboard/report decision

Reporting must provide both:

- filtered tables for exact values and operational review,
- charts/trends for pattern recognition, comparison and change over time.

Tables are the primary numeric surface; charts are complementary. The same filters and finalized historical truth must drive both so chart values always reconcile with the underlying table rows.

Normal users access reports through application UI, not raw technical worksheets.

## Smallest durable logical schema now drafted

`Documentation/STATISTICS_HISTORY_DESIGN.md` defines the minimum logical fields for:

- finalized day snapshot headers,
- therapist/day workload rows,
- append-only event rows.

The weekly therapist workload view will derive from finalized therapist/day rows, with actual final occupied timeslots as the primary metric and final active sessions as a separate workload metric.

## Statistics/history design checkpoint status

The core architecture questions are now complete enough to stop abstract questioning and collect the user's concrete requirements.

Remaining details should be derived from those requirements:

- exact weekly/monthly comparison presets,
- additional metrics worth persisting versus deriving,
- exact patient-level statistics/reports,
- dashboard views and drill-downs that map to real operational use.

## Exact resume point

1. Treat the `27 passed in 3.09s` presentation gate as green.
2. Treat explicit `Κλείσιμο ημέρας` finalization as locked.
3. Treat the hybrid persistence model as locked: finalized daily snapshots + append-only events.
4. Treat auditable versioned correction through `Διόρθωση κλεισμένης ημέρας` as locked.
5. Treat same-workbook hidden/protected technical storage as locked.
6. Treat indefinite retention by default as locked.
7. Treat patient-level historical analysis as required.
8. Treat both tables and charts/trends as required for reporting.
9. Collect the user's own requested history/statistics now, before implementation.
10. Use those requirements to finalize the first concrete slice: day finalization + therapist/day history + weekly workload view.
11. Do not manually save smoke workbooks outside a tested flow.
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
