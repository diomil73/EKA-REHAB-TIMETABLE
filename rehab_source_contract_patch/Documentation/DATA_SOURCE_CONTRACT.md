# Data Source Contract

This document freezes the source-of-truth rules before Python is allowed to write to Excel.

## Authoritative workbook sources

| Sheet | Role |
|---|---|
| `PATIENTS` | Patient identity, room, infectious flag and current status |
| `PATIENT_PLANNER` | Recurring base programme |
| `SETTINGS` | Existing configuration values, timeslots, day patterns and allowed lists |

Python may read these sheets into an immutable in-memory snapshot. The read-only adapter must never save the workbook.

## Explicitly not authoritative

- `SESSIONS`: legacy/audit-only while PatientID/name drift exists.
- `MASTER_SCHEDULE`: derived human-readable view.
- `THERAPIST_DAILY`: derived daily therapist view.
- `CONFLICT_LOG`: derived conflict output.
- `REPLACEMENTS`: derived suggestions/selection UI.
- `STATISTICS`: derived report.
- `NEW_PATIENT`: UI only; committed patient data belongs in `PATIENTS`.

## Legacy operational sheets pending redesign

`THERAPIST_ATTENDANCE`, `REPLACEMENT_LOG` and `DAILY_LOG` remain part of the existing workbook workflow, but Python does not silently treat them as authoritative sources yet. Their final ownership will be decided before write integration.

## Safety gates

A normal operational snapshot is blocked when an authoritative-source error exists, including:

- duplicate `PatientID` values in `PATIENTS`,
- PatientID/name disagreement between `PATIENTS` and `PATIENT_PLANNER`,
- unsupported day patterns.

Warnings do not get guessed away. For example, a timed entry with no day pattern remains unresolved and is omitted from materialized schedule data until an explicit rule is confirmed.

## Privacy rule

The snapshot stays in memory. The tooling does not automatically export patient names or patient schedules to JSON/CSV files that could accidentally be committed to GitHub.
