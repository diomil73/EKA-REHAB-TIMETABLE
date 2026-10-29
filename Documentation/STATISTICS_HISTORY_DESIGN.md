# Statistics & History Design

Last updated: 2026-10-24

## Purpose

This document is the specialist source of truth for historical facts, therapist workload history, statistics and user-facing reporting.

The design principle is that historical metrics must come from persisted historical facts/events or finalized-day snapshots, never by re-reading today's mutable workbook projection as if it represented the past.

## Primary therapist workload metric

The primary historical workload metric is the **actual final operational workload of the day** after concrete-date absences, cancellations/no-shows, replacements, suspensions and approved overrides have been applied.

For each therapist/day, persist at least:

- final occupied timeslots,
- final active/performed sessions,
- replacements received,
- cancellations/no-shows that changed the final operational workload,
- treatment suspensions affecting the final day,
- explicit approved double-booking/over-capacity exceptions when they occur.

Occupied timeslots and active sessions are separate measures. A legitimate shared slot may contain two sessions but consume one occupied timeslot.

### Reporting scope decision: unused/idle capacity

Do **not** include `available/unused timeslots`, `unused capacity`, `idle time`, or equivalent "unexploited available time" metrics in the primary therapist reporting scope.

The reporting system should focus on actual work performed and final operational outcomes, not on interpreting unoccupied capacity as a staff-performance measure.

This metric is deliberately set aside unless the user explicitly reopens it in a future design decision.

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

## Physical storage decision

Historical snapshots, therapist/day workload rows and append-only events stay inside the same macro-enabled workbook as hidden/protected technical tables.

Do not make normal operation depend on external history files. This keeps the workbook and its historical truth together and reduces the risk of orphaned or mismatched history files.

Normal users must never need to navigate to these technical tables. Access is through application buttons/forms/reports only. Any future export is a deliberate reporting/administrative action, not the authoritative store.

The existing privacy rule remains: do not automatically export patient-level history into repository files or other uncontrolled external artifacts.

## Smallest durable logical schema

The first implementation should keep the logical schema narrow and stable. Physical Excel sheet/table names remain an implementation detail.

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

## Patient-level analysis

Patient-level historical analysis is a required product capability, not merely an aggregate reporting option.

### Patient presence in the system

For every patient retain the date on which the patient first entered the operational table/system. From that date the application derives and displays the elapsed stay as calendar **years, months and days**.

The entry date is a historical fact and must not be replaced by a later schedule edit. If a future discharge/re-entry workflow introduces separate care episodes, episode semantics must be designed explicitly rather than silently changing the meaning of the original entry date.

### Treatment-day and lost-treatment history

For each patient the application must derive/display:

- treatment/service days,
- scheduled treatment sessions,
- performed treatment sessions,
- total lost/missed treatment sessions,
- the reason for every lost/missed session,
- totals and trends by period.

Weekends and official holidays are excluded from expected-working-day calculations. Official holidays therefore need a reliable calendar/configuration source; the system must not infer them from missing sessions.

A lost treatment remains session-level historical truth. Reasons must remain structured enough to distinguish at least patient absence/no-show, department cancellation/postponement, therapist-related disruption where applicable, treatment suspension and any other agreed operational reason.

### Treatment portfolio and timeline

For each patient retain/derive the number of **different treatment types** received and display the treatments on a running chronological timeline. The timeline should make treatment starts, stops and meaningful changes visible over the patient's stay rather than showing only today's programme.

### 15-day progress/evaluation history

The `MASTER` physiotherapy-time cell is a clickable application action that opens the patient's progress/evaluation view.

The progress view is completed manually approximately every **15 days** and is historical/append-only. It uses explicit milestones. Each milestone row contains:

- the examined question/goal,
- `Από` date,
- `Έως` date,
- a manual outcome for evolution during that period.

Locked outcome semantics from the current requirement:

- 🟢 positive/progress,
- ⚪/grey neutral or no meaningful progress,
- 🔴 negative deterioration.

The exact icon artwork may be refined in the final UI, but green/grey/red meaning is fixed. The prior roadmap idea of a yellow neutral state is superseded by the user's current grey-neutral requirement.

The patient progress history must support a chronological trend/sparkline style view and must never be detached from the patient by later therapist/time/cell changes. Internally it is keyed by stable `PatientID` and, where useful, period/session/date identifiers.

The older UI roadmap also records a useful aggregate view of current inpatients with room, patient name, current 15-day goal, latest evaluation, historical evaluation sequence and a small trend visual. It also proposes a high-level department/hospital trajectory indicator, but that aggregate must not be implemented as a naive average; its clinical meaning needs an explicit rule before use.

### Intranet patient-folder link

The visible patient name in `MASTER` is a clickable application action that opens the patient's electronic folder on the hospital intranet/shared network location.

The older recorded architecture is retained:

- resolve the folder from stable `PatientID` / `FolderKey`, not from the visible patient name or cell address,
- prefer an identity-based UNC/shared-network path over a mapped drive letter,
- keep internal network-root configuration outside GitHub when it contains infrastructure details,
- respect Windows/intranet permissions; the workbook must not bypass them,
- moving/renaming timetable cells must not break the folder link.

### Additional patient views already recorded in earlier roadmap notes

Retain these earlier ideas as planned report/list capabilities:

- patients with an overdue 15-day evaluation,
- lost treatments by selected period,
- search/open actions leading to progress history, intranet folder and operational patient information,
- scheduled vs performed, absences/no-shows, cancellations/postponements, suspensions, replacements with effective therapist/time and permanent therapist-assignment changes.

Patient-level reports use stable `PatientID` internally while displaying normal patient identity/business labels in the UI. Historical truth comes from finalized snapshots/events, not the patient's current recurring schedule.

Patient-identifiable analysis remains inside the controlled application/workbook environment by default. Technical history sheets stay hidden/protected, and no automatic repository export of patient-level data is permitted.

## Therapist-level analysis

Therapist-level analysis is also a required product capability.

### Workload history

Retain/display daily, weekly, monthly and longer-term workload history from finalized therapist/day facts. The personal therapist card should make persistent overload, high-load periods, lower-load periods and long-term trend visible.

### Calendar-year attendance and absence history

For each therapist, retain/display presence/attendance within the selected calendar year and a dated absence history for the same year.

Every therapist absence must carry a structured reason. The current required reason catalogue is:

- `Κανονική άδεια`,
- `Ρεπό`,
- `Αναρρωτική`,
- `Αιμοδοσίας`,
- `Του νόμου`,
- `Εκπαιδευτική`,
- `Τέκνων`,
- `Ειδική`,
- `Από τη σημαία`.

These labels are business-facing requirements. Internal reason codes may be stable technical values, but the UI must show the agreed Greek wording.

The attendance report should reconcile presence days/hours, absence days/hours and reason breakdown rather than maintaining a manually typed annual total.

### Additional therapist statistics retained in scope

The therapist statistical set remains focused on actual work and operational outcomes:

- total treatments,
- daily/weekly/monthly workload,
- replacements accepted,
- infectious-patient workload,
- robotic workload,
- cancellations and patient absences affecting the final operational result,
- sessions lost because of therapist absence,
- sessions successfully covered through replacement,
- absence dates/counts/reasons and related operational replacement history,
- current active-patient/load view,
- quick list of therapists with high workload.

`Available/unused timeslots` and equivalent idle-capacity metrics are intentionally excluded from the current reporting scope unless explicitly reopened later.

Therapist statistics should be derived from assignments, finalized daily facts and operational event history wherever possible, not maintained as parallel manual totals.

## Department-level analysis

Every core operational metric available for an individual therapist must also be available as a **department-wide aggregate** for the same selected period, using the same finalized snapshots and event history.

The department total is not a separately maintained dataset. It is derived from the same therapist/day and session-level finalized facts so individual and departmental numbers always reconcile.

Department reporting should support at least:

- total treatments/sessions performed,
- total occupied therapist timeslots as an operational workload measure,
- total replacements accepted/performed across the department,
- infectious-patient workload,
- robotic-treatment workload,
- patient absences/no-shows and department cancellations/postponements,
- sessions affected by therapist absence,
- sessions successfully covered through replacement,
- treatment suspensions,
- approved double-booking/over-capacity exceptions,
- therapist attendance/absence totals and structured reason breakdown where appropriate,
- daily, weekly, monthly and longer-period trends.

The reporting UI must allow a user to move between **individual therapist** and **whole department** views without changing the underlying metric definitions. A department figure should be drillable back to the therapists/days/events that compose it.

Unused/idle-capacity metrics remain excluded at department level as well, unless explicitly reopened later.

## Dashboard and report UX

The reporting UI must support **both filtered tables and charts/trends**.

The default interaction pattern is:

- a clear table/report surface first, with filters for period, therapist, patient and relevant status/category,
- charts and trend visuals as a complementary layer for pattern recognition, comparison and change over time,
- drill-down from aggregate/graph views to the underlying finalized historical records where appropriate,
- no requirement for users to navigate raw Excel data sheets.

Tables remain the primary source for exact operational values. Charts are used to expose trends, workload balance, planned-versus-actual gaps, absence/no-show patterns, replacement load and patient-level trajectories without hiding the underlying numbers.

Weekly and monthly reporting should use the same historical truth and filtering model so that a value seen in a chart can be reconciled with the corresponding table rows.

## Retention and archive policy

Historical snapshots and append-only events are retained indefinitely by default for the life of the system.

The application must not automatically delete, truncate or age out historical records. If archive or data-reduction functionality is introduced later, it must be an explicit controlled administrative action with validation and auditability. Ordinary reporting should continue to work across all retained finalized history.

This policy applies to both current-final snapshots and superseded historical versions needed for amendment/audit chains.

## Open design questions

The core architecture is now locked. Remaining details should be driven by concrete operational requirements and implementation discovery:

1. exact weekly/monthly comparison presets and date ranges,
2. authoritative official-holiday calendar/configuration,
3. detailed therapist presence semantics for partial-day attendance/absence when that case is introduced,
4. final milestone-row fields and interaction details for the 15-day progress form,
5. exact intranet `FolderKey`/path-resolution configuration,
6. additional patient/therapist statistics the user remembers later,
7. which derived metrics deserve snapshot columns for performance versus calculation on demand.

## Relationship to current scheduling truth

Current operational truth remains:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics/views`

Historical workload must capture the final effective result of that concrete-date truth rather than recomputing history from later recurring schedules.
