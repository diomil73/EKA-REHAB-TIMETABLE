# Statistics & History Design

Last updated: 2026-10-20

## Purpose

This document is the specialist source of truth for historical facts, therapist workload history, statistics and user-facing reporting.

The design principle is that historical metrics must come from persisted historical facts/events or finalized-day snapshots, never by re-reading today's mutable workbook projection as if it represented the past.

## Primary therapist workload metric

The primary historical workload metric is the **actual final operational workload of the day** after concrete-date absences, cancellations/no-shows, replacements, suspensions and approved overrides have been applied.

For each therapist/day, persist at least:

- final occupied timeslots,
- final active/performed sessions,
- replacements received,
- cancellations/no-shows that released capacity,
- treatment suspensions affecting the final day,
- explicit approved double-booking/over-capacity exceptions when they occur.

Occupied timeslots and active sessions are separate measures. A legitimate shared slot may contain two sessions but consume one occupied timeslot.

## Secondary planned-vs-actual context

The initially planned workload may also be retained as a secondary comparison dimension, but it is not the primary user-facing measure.

This supports questions such as:

- planned 6/6, actual 4/6 after two cancellations,
- whether a therapist repeatedly absorbs replacement work,
- whether recurring planning systematically over- or under-estimates operational load.

The default dashboard/history view should emphasize **actual final workload first**. Planned workload is supporting context and can be shown in drill-down or comparison views.

## Weekly therapist history

The weekly view should aggregate daily final workload for each therapist, typically Monday-Friday, and expose at minimum:

- occupied timeslots per day, e.g. `4/6`, `5/6`, `6/6`,
- active/performed sessions per day,
- weekly total sessions,
- average occupied timeslots,
- count of full-capacity days,
- replacements received during the week,
- approved over-capacity/double-booking events,
- optional planned-vs-actual comparison.

Normal capacity semantics remain:

- 0-4/6: green,
- 5/6: yellow/orange,
- 6/6 or above: red.

## Finalization semantics

Historical daily workload becomes final only through an explicit application action: **`Κλείσιμο ημέρας`**.

The application must not silently finalize a day just because the clock passed a certain hour or the workbook was closed. The user deliberately closes the operational day when the day's exceptions have been resolved.

Before `Κλείσιμο ημέρας` can succeed, the daily state must pass the finalization gate:

- no displaced/orphan session remains unresolved,
- every concrete-date session has a final operational outcome,
- replacements, cancellations/no-shows and suspensions are already reflected in the effective daily truth,
- workload/statistics are computed from that final effective state.

On successful close, persist the finalized daily historical snapshot. At minimum this must include the therapist/day workload facts needed for weekly history and the final daily outcome counts.

## Post-finalization correction workflow

A closed day is not reopened and destructively rewritten as if the first close never happened.

The application must expose a deliberate action such as **`Διόρθωση κλεισμένης ημέρας`**.

A correction must:

1. require explicit selection of the already-finalized service date,
2. record an amendment event describing that a historical correction was started,
3. apply the corrected operational facts through the same validated business rules used by the normal daily workflow,
4. require the finalization gate to pass again before the correction is committed as final,
5. create a new finalized snapshot version,
6. preserve the previous snapshot version as historical/superseded rather than deleting it,
7. link the new snapshot to the snapshot it supersedes,
8. record who/what performed the correction and when,
9. keep weekly/monthly dashboards pointed at the latest current-final snapshot while audit views can show the full version chain.

A correction that fails validation must not partially replace the current-final snapshot.

This makes the historical record versioned and auditable without forcing ordinary users to work directly with technical revision IDs.

## Persistence model

The historical architecture is **hybrid**.

Use both:

1. **Finalized daily snapshots** for fast operational reporting, weekly/monthly aggregation and stable historical totals.
2. **Append-only historical events** for auditability, corrections and reconstruction of how the final result was reached.

The snapshot is the reporting surface, not the only historical truth. Events must not be destructively overwritten when a later correction is needed.

The minimum event stream should be able to represent meaningful day-changing actions such as:

- therapist/patient absence,
- cancellation or no-show,
- replacement assignment,
- treatment suspension,
- explicit approved double-booking/over-capacity exception,
- permanent therapist-assignment change when it affects historical interpretation,
- day finalization,
- post-finalization amendment/correction.

Weekly and monthly dashboards should normally read finalized snapshots. Drill-down/audit views may follow the corresponding event history.

## Smallest durable logical schema

The first implementation should keep the logical schema narrow and stable. Physical Excel sheet/table names remain an implementation detail until the persistence location is selected.

### Finalized day snapshot header

One snapshot header per service date/version:

- `snapshot_id`: immutable internal ID,
- `service_date`: operational date,
- `version`: monotonically increasing revision number for that date,
- `status`: at least `CURRENT_FINAL` or `SUPERSEDED`,
- `finalized_at`: timestamp of the successful close/correction finalization,
- `supersedes_snapshot_id`: previous finalized version when this is an amendment, otherwise empty,
- `unresolved_sessions`: must be `0` for a finalized snapshot,
- `planned_session_count`: optional secondary planned context,
- `final_session_count`: total final active/performed sessions for the day,
- `patient_absence_count`,
- `cancellation_count`,
- `no_show_count`,
- `replacement_count`,
- `treatment_suspension_count`,
- `approved_exception_count`: explicit double-booking/over-capacity approvals.

### Therapist/day workload rows

One row per therapist within a snapshot:

- `snapshot_id`,
- `service_date`,
- `therapist_id`,
- `capacity_timeslots`,
- `planned_occupied_timeslots`: secondary context,
- `planned_sessions`: secondary context,
- `final_occupied_timeslots`: primary capacity metric,
- `final_active_sessions`: primary patient/session workload metric,
- `replacements_received`,
- `released_by_cancellation_or_no_show`,
- `treatment_suspensions_affecting_load`,
- `approved_shared_or_over_capacity_events`.

The weekly history is derived from these finalized therapist/day rows, never from later recurring schedules.

### Append-only event rows

Each historical event should contain at least:

- `event_id`: immutable internal ID,
- `service_date`,
- `event_type`,
- `recorded_at`,
- `actor_or_source`: user/workflow/system source identifier,
- `therapist_id`: optional when relevant,
- `patient_id`: optional stable internal identifier when relevant,
- `session_id`: optional when a concrete session is involved,
- `reason_code`: structured reason when applicable,
- `related_event_id` or `related_snapshot_id`: link for amendment/finalization chains,
- `summary`: concise auditable description without relying on mutable display text.

Event rows are append-only. Corrections create new events; they do not edit the original event into a different historical claim.

## Open design questions

Still to lock before implementation:

1. physical storage location for snapshots/events,
2. retention horizon and archive policy,
3. weekly/monthly comparison windows,
4. patient-level historical statistics and privacy boundaries,
5. dashboard layout and drill-down UX,
6. which statistics are operationally useful enough to keep permanently versus derive on demand.

## Relationship to current scheduling truth

Current operational truth remains:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Historical workload must capture the final effective result of that concrete-date truth rather than recomputing history from later recurring schedules.
