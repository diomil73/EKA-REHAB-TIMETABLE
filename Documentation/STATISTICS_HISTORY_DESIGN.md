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

A later correction to historical data must be auditable. Do not silently rewrite historical truth without retaining the fact that a correction occurred. The exact reopen/amendment workflow remains to be designed.

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

## Open design questions

Still to lock before implementation:

1. exact finalized-day snapshot schema and storage location,
2. exact append-only event schema and identifiers,
3. correction/reopen/amendment workflow after a day is finalized,
4. retention horizon and archive policy,
5. weekly/monthly comparison windows,
6. patient-level historical statistics and privacy boundaries,
7. dashboard layout and drill-down UX,
8. which statistics are operationally useful enough to keep permanently versus derive on demand.

## Relationship to current scheduling truth

Current operational truth remains:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Historical workload must capture the final effective result of that concrete-date truth rather than recomputing history from later recurring schedules.
