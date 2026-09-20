# Sequential therapist-absence replacements

When one therapist absence affects several patients, replacement ranking must not be a static report.

The confirmed workflow is sequential:

1. Read DAILY_INPUT absences.
2. Exclude sessions whose patient is absent.
3. Show ranked providers for the first affected session.
4. After the user accepts one provider/time, add that replacement to the operational overlay.
5. Recalculate provider load and free timeslots before showing the next patient.
6. Continue until every affected session is assigned or explicitly left unresolved.
7. Build a new patient-centric `.xlsm` preview. Never overwrite the input or baseline workbook.

This prevents one low-load therapist from appearing with the same old load for every patient in a multi-patient absence event.
