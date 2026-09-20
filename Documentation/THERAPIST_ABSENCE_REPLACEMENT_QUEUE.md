# Therapist absence replacement queue

`DAILY_INPUT` remains the human input surface. A therapist absence is expanded into
individual patient sessions for the selected date. Each patient assignment remains
independent, including members of visual pairs/groups.

Rules:

- Patient absence has priority. If the patient is also absent, no replacement is needed.
- A therapist absence affects only session occurrences covered by its date/time window.
- Every affected patient gets an independent ranked replacement list.
- Another therapist declared absent in `DAILY_INPUT` cannot be suggested.
- Ranking remains load-first. At equal load, exact-time availability is preferred.
- A busy exact time does not exclude a provider if another common timeslot is available.
- This planning step is read-only. Acceptance/application is a later workflow step.
