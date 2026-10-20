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

Historical daily workload should be treated as final only when the operational day is finalized, meaning no displaced session remains unresolved. The precise persistence trigger and any correction/amendment workflow remain to be designed.

A later correction to historical data must be auditable. Do not silently rewrite historical truth without retaining the fact that a correction occurred.

## Open design questions

Still to lock before implementation:

1. exact finalized-day snapshot schema and storage location,
2. whether history is append-only events, daily snapshots, or a hybrid of both,
3. correction/amendment workflow after a day is finalized,
4. retention horizon and archive policy,
5. weekly/monthly comparison windows,
6. patient-level historical statistics and privacy boundaries,
7. dashboard layout and drill-down UX,
8. which statistics are operationally useful enough to keep permanently versus derive on demand.

## Relationship to current scheduling truth

Current operational truth remains:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Historical workload must capture the final effective result of that concrete-date truth rather than recomputing history from later recurring schedules.
