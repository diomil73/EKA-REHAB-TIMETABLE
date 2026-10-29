# Patient lifecycle dates

Last updated: 2026-10-24

## Purpose

Record the product and data contract for patient start/end dates, including migration of the existing active patient population and compact MASTER presentation.

## Core lifecycle fields

Every patient must have two explicit historical lifecycle dates:

- `StartDate`: date the patient entered the rehabilitation operational programme.
- `EndDate`: date the patient left the active programme and was archived.

These are historical patient facts and must not be inferred later from the current recurring timetable.

## New patient registration

When a new patient is registered, `StartDate` is populated automatically from the registration/operational start date. The default may be the current service date, but the registration workflow must allow the operator to correct it before final confirmation when the true start date differs.

`EndDate` remains empty while the patient is active.

## Archive workflow

When an active patient is moved to archive, the archive workflow records `EndDate` automatically from the archive/discharge date.

The operator must be able to correct the proposed end date before final confirmation when the real last operational day differs from the action date.

Archiving must not delete or overwrite `StartDate`.

## Manual correction / migration requirement

Both `StartDate` and `EndDate` must be editable through a controlled application form, not by requiring normal users to edit raw technical cells.

This is required immediately because approximately 95 patients are already present in the programme and their historical start dates predate the new lifecycle-date feature.

Migration rule:

- existing active patients may have `StartDate` entered manually;
- archived/historical patients may have both `StartDate` and `EndDate` corrected manually when known;
- manual corrections should be auditable once the historical event model is active;
- `EndDate` cannot be earlier than `StartDate`;
- an active patient normally has an empty `EndDate`;
- an archived patient should normally have a populated `EndDate`.

## MASTER presentation

The active patient's start date should be visible directly in `MASTER_SCHEDULE`, positioned near the responsible-doctor display if layout permits.

Presentation must be compact so it does not widen the operational table unnecessarily.

Preferred compact label:

`Έν. 12/09/26`

Alternative if space is tighter:

`Εν:12/09/26`

Use `dd/mm/yy` in the MASTER display. The underlying stored value remains a true date value, not preformatted text.

The end date is not normally shown in the active MASTER row because active patients have no end date. It belongs in archive/history views and the patient card.

## Reporting relationship

Patient stay duration is derived from these lifecycle facts:

- active patient: `StartDate -> current selected/report date`;
- archived patient: `StartDate -> EndDate`.

The patient-history UI may display elapsed stay as years, months and days, as already defined in the statistics/history design.

Department-wide and clinic-level patient reporting must use the same lifecycle dates for counts such as active patients, admissions/starts, archived/discharged patients and length-of-stay analyses.

## Safety and identity

Lifecycle dates belong to the stable patient identity (`PatientID`), never to a MASTER row number or current cell address. Moving a patient within the timetable must not change these dates.

Any later care-episode/re-admission model must be designed explicitly rather than silently overwriting the original lifecycle history.
