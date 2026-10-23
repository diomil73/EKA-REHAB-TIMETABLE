# Current Handoff

Last updated: 2026-10-24

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
- double-booking review/install/apply focused gate: `13 passed in 1.63s`
- double-booking UI smoke V5: native green; confirmed double-booking persists in-place without close/reopen and projects immediately into `THERAPIST_DAILY`
- global replacement / reclining / service-policy focused gate: `38 passed in 5.30s`
- student production wiring focused gate: `39 passed in 2.35s`

Restore performance is accepted for now and deferred until end-stage polish.

## Current milestone: student replacement production wiring closed

The double-booking exception workflow is operationally closed end-to-end.

The global replacement path already has reclining adjacency and special pool/robotic service hours wired into production. The policy layer also preserves student identity as `STUDENT` rather than treating students as therapists.

The remaining student production gap was that `Python/tools/apply_global_daily_replacements.py` did not load the authoritative `STUDENTS` registry and did not pass students through the queue, global solver and `THERAPIST_DAILY` patient-centric projection.

That wiring is now implemented and verified:

- commit `c8e571d` loads authoritative students with `read_students(input_book)`, passes them into the therapist-absence queue, global replacement solver and patient-centric projection, and safety-stops on `StudentRegistryError`;
- commit `836e610` adds a regression guard that requires the production tool to keep this full student path wired;
- local Windows focused gate after the wiring: `39 passed in 2.35s`.

Treat authoritative student candidate loading + solver wiring + `THERAPIST_DAILY` student projection as closed unless a later native smoke reveals a display/layout defect.

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

## Concrete patient-history requirements now locked

For each patient, preserve/derive:

- date first entered the operational table/system,
- elapsed stay shown as years, months and days,
- treatment/service days,
- scheduled and performed sessions,
- total lost/missed treatments with structured reason,
- expected-working-day logic excluding weekends and official holidays,
- count of different treatment types,
- chronological treatment timeline showing starts/stops/meaningful changes,
- full patient-level trends/totals over selectable periods.

`MASTER` patient links:

- patient **name** opens the intranet/shared-folder patient record using stable identity-based folder resolution,
- patient **physiotherapy time** opens the 15-day progress/evaluation view.

The 15-day progress view is manual, append-only and milestone-based. Each milestone row carries the examined question/goal, `Από`, `Έως` and one outcome: green positive/progress, grey neutral/no meaningful progress, red negative/deterioration. The previous yellow-neutral roadmap idea is superseded by the current grey-neutral requirement.

Older roadmap notes also retain the planned current-inpatient progress overview, trend/sparkline, overdue-15-day-evaluation list and optional department-level aggregate trajectory, provided the aggregate later gets a clinically meaningful rule rather than a naive average.

## Concrete therapist-history requirements now locked

For each therapist, preserve/derive:

- workload history,
- calendar-year attendance/presence,
- calendar-year dated absences,
- structured absence reason.

Required user-facing absence reasons:

- `Κανονική άδεια`,
- `Ρεπό`,
- `Αναρρωτική`,
- `Αιμοδοσίας`,
- `Του νόμου`,
- `Εκπαιδευτική`,
- `Τέκνων`,
- `Ειδική`,
- `Από τη σημαία`.

Earlier project notes also already proposed therapist statistics for total treatments, daily/weekly/monthly workload, replacements accepted, infectious-patient workload, robotic workload, available/unused timeslots, cancellations/patient absences affecting utilization, sessions lost due to therapist absence and sessions covered by replacement. Keep these in the planned metric set unless explicitly removed later.

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
- append-only event rows,
- patient-level history/timeline requirements,
- therapist attendance/absence history requirements.

The weekly therapist workload view will derive from finalized therapist/day rows, with actual final occupied timeslots as the primary metric and final active sessions as a separate workload metric.

## Statistics/history checkpoint status

The core architecture questions are complete. Concrete patient/therapist reporting requirements are recorded and remain locked while the replacement-quality vertical is finished.

Remaining implementation-design details:

- official-holiday calendar/configuration source,
- partial-day therapist attendance/absence semantics if needed,
- exact 15-day milestone form layout/fields,
- intranet `FolderKey`/path-resolution configuration,
- any additional user-requested metrics remembered later,
- which derived metrics deserve snapshot columns for performance versus calculation on demand.

## Exact resume point

1. Student production wiring is green: `39 passed in 2.35s`.
2. Treat authoritative student candidate loading + solver wiring + `THERAPIST_DAILY` student projection as closed unless a later native smoke reveals a display/layout defect.
3. Do not reopen the already-closed double-booking no-close/no-reopen workflow unless a regression appears.
4. Reassess the remaining replacement-quality checklist before opening the next implementation vertical.
5. Keep the statistics/history decisions below locked for the later day-finalization/history slice.
6. Do not manually save smoke workbooks outside a tested flow.
7. Do not return to restore performance tuning yet.

## Open replacement-quality milestone

The replacement lifecycle is correct. The global assignment path now includes reclining adjacency, special pool/robotic service hours and student provider identity/wiring, but the broader replacement-quality checklist remains the controlling contract until explicitly closed.

All known therapist absences for a date must be entered first. Then all displaced sessions are solved together, not greedily.

Locked highlights:

- maximize placed displaced patients and minimize unresolved orphans,
- protect unique viable slots from flexible patients,
- if A can use 09:00/10:00/11:00 and B only 09:00, protect 09:00 for B,
- prefer original time when it does not damage a higher-priority constraint,
- keep ΦΘ and reclining treatment close where feasible,
- keep workload balanced and avoid a two-timeslot imbalance where feasible,
- 13:00 is only a weak soft preference to avoid,
- pool/robotic operating hours must remain persistent configuration,
- students remain distinct `STUDENT` providers and must be sourced from the authoritative `STUDENTS` registry,
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

Design/product ideas requested by the user must be recorded in `Documentation/FUTURE_FEATURES.md` or the appropriate specialist roadmap/decision document when they are agreed, so they remain available when their implementation phase arrives.
