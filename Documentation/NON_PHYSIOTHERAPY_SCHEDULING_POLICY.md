# Non-Physiotherapy Scheduling Policy

Status: agreed architecture contract, implementation pending.

## Purpose

Occupational Therapy, Speech Therapy, Psychology and Special Physical Education do not use the Physiotherapy workload/replacement engine as their normal scheduling model.

These specialties have lower and different workload patterns, including legitimate cases where one therapist may have multiple patients at the same date/time.

Their scheduling model is therefore deliberately simpler:

`manual specialty scheduling + canonical time selection + cross-specialty patient conflict guard`

## Applies to

This policy applies to:

- `OCCUPATIONAL_THERAPY`
- `SPEECH_THERAPY`
- `PSYCHOLOGY`
- `SPECIAL_PHYSICAL_EDUCATION`

It does not replace the dedicated Physiotherapy scheduling model.

## Manual scheduling ownership

Each specialty controls its own schedule within its authorization scope.

- `PARAMEDICAL_HEAD` may edit all non-Physiotherapy specialties.
- `DEPARTMENT_HEAD` may edit only the specialty assigned to that user.
- `SUPER_USER` may edit all specialties.

The system assists and validates; it does not automatically optimize these specialties in the same way as Physiotherapy.

## Canonical time rule

All appointment times must be selected from the application's authoritative time lists.

Users must not be allowed to type arbitrary variants of a time or bypass the time list with an override.

Examples of forbidden behavior include:

- free-text times;
- alternate formatting for the same time;
- ad-hoc times outside the canonical list;
- user override of the canonical-time validation.

This is a hard data-integrity rule because cross-specialty conflict detection depends on normalized times.

## Same-specialty multiple bookings

The system must not impose Physiotherapy-style one-patient-per-therapist-per-slot capacity on these specialties.

A specialty may legitimately place multiple patients with the same therapist at the same date/time when that matches the specialty's working model.

Therefore:

- same therapist + same date/time + multiple patients can be legal;
- each specialty remains responsible for its own internal workload/capacity decisions;
- the shared scheduling guard must not reject a booking merely because another patient of the same specialty already uses that therapist/time.

## Cross-specialty patient conflict rule

The shared system must prevent a patient from being scheduled at a time already occupied by a different specialty on the same day.

Before accepting a non-Physiotherapy booking, the system must evaluate the selected patient's effective schedule for that date.

If the selected time is already occupied by another specialty:

1. the booking is not accepted at that time;
2. the operator sees an operational message identifying the conflicting specialty;
3. the system computes the patient's free canonical times for that day;
4. those free times are offered for selection;
5. there is no user override that can force the cross-specialty conflict through.

Example operator meaning:

`Η ώρα 10:30 είναι ήδη κατειλημμένη από Λογοθεραπεία. Επιλέξτε μία από τις διαθέσιμες ώρες του ασθενή.`

The exact wording/UI may change, but the blocking behavior is mandatory.

## Conflict ownership principle

The global constraint is the patient's availability across specialties, not a universal therapist-capacity rule.

This means:

- the application coordinates the patient's timetable across departments;
- each non-Physiotherapy specialty owns its own therapist-capacity practice;
- legitimate same-specialty group/concurrent work is preserved;
- cross-specialty patient overlaps are blocked.

## Workload / statistics / restore boundary

Physiotherapy workload, management statistics and automated daily-change/restore tooling are not general features for the other specialties.

Access to those Physiotherapy supervisory functions is limited to:

- `SUPER_USER`; and
- any separately authorized `PHYSIOTHERAPY_HEAD` appointed in the future.

Other specialties perform their own schedule changes manually under the validation rules in this document.

## Implementation direction

The implementation should expose a reusable validation service, conceptually:

`validate_specialty_booking(patient, date, specialty, canonical_time)`

Its responsibilities are:

- validate that the submitted time is canonical;
- determine the effective patient schedule for the selected date;
- ignore same-specialty therapist-capacity assumptions;
- detect conflicts owned by another specialty;
- return the conflicting specialty when blocked;
- return the patient's free canonical times for the day.

The UI may be Excel/VBA, but the authoritative validation belongs in the Python business core.

## No override rule

No normal operator role may bypass:

- canonical-time validation; or
- cross-specialty patient-conflict validation.

If a future exceptional administrative correction path is ever required, it must be designed as a separate audited workflow rather than a generic `Override` button.
