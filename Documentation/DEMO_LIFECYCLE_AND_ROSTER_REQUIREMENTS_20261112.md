# Diomil functional demo: outpatient end, inpatient permanent discharge, therapist/student roster lifecycle

**Operator instruction (2026-11-12):** Pilot is synthetic-data-only; actual 2027 program NOT imported. Operator enters outpatient appointments through the existing outpatient management menu. Cosmetic finish is postponed. All patient, roster and daily logic must be proven in one functional demo before clinical use.

## Required behavior

1. **Outpatient program ends:** Make an explicit end-program action reachable from the existing outpatient schedules menu. Effective date is recorded; stop generating recurring outpatient bookings after that date (inclusive end date means last allowed treatment date). The patient must no longer occupy future daily timeslots after the stop. Preserve prior/closed finalized sessions, original program and audit trail. Read-only preview of affected upcoming bookings before committing. Do **not** equate program end with an absence/cancellation for an individual day.
2. **Inpatient permanently discharged:** Make an explicit permanent discharge workflow via the patient management UI. Record effective discharge date; remove patient from active future inpatient roster/timeslot generation from the next applicable day, without deleting patient identity or changing historical finalized visits, absent reasons or records. Protect against reuse of formerly occupied appointments. Existing statuses/workflows must be reused where applicable; do not introduce new patient absence types. Return/re-admission, if needed, must be an explicit separately validated action, not silent reactivation.
3. **Add and remove therapists:** Use one managed therapist roster, with add, edit and **deactivate/retire (not hard-delete)** actions. Newly active therapist must appear in all relevant ΦΘ/robotics selectors and planning capacity. Retirement must block if the therapist still owns future appointments unless they are explicitly transferred or resolved. Preserve stable historical therapist IDs, names and archived work.
4. **Add and remove students:** Use the separate `STUDENTS` registry. Student onboarding requires StudentID, name, start/end placement dates and supervisor where applicable. Student end/withdraw/deactivation must remove active choices after the effective date, identify affected recurring assignments, and drive safe takeover/reassignment. Preserve identity/history. Existing `student_transition` is a **read-only suggestions planner**, not proof of a committed end-placement workflow.
5. **Daily consistency:** Active census, capacity, `MASTER_SCHEDULE`, `THERAPIST_DAILY`, outpatient menus and statistics must reflect the same effective schedule. One shared slot counts once for capacity. No finalized historical rewriting from a roster change. All writes only through validated native Excel/COM transaction with rollback/reopen/read-back as applicable.

## Existing code inventory inspected

- Outpatient schedule UI: `Python/rehab_excel/outpatient_schedule_manager_vba.py` presents listed schedules, create and edit actions. **No end-program action found** in the inspected manager source.
- Student registry: `Python/rehab_excel/student_registry.py`, `student_registry_schema.py`, `student_registration_form_vba.py`. `Python/rehab_core/student_transition.py` provides read-only end/hand-off suggestions but **does not itself write a permanent transition**.
- Patient state: `Python/rehab_core/models.py` has existing patient types and statuses, which must be consulted rather than multiplied.
- Staff availability/therapist registry: code exists across scheduling and settings; a verified manager for safe add/retire has **not been established** by this review. Do not claim removal already complete.
- Operator enters real outpatients through existing menu **after go-live**, not by importing the external-patient sheet of the uploaded workbook.

## Acceptance tests to implement before demo claim

- Ending an outpatient program leaves past history, eliminates future appointments, frees impacted slot and prevents reappearance on refresh/reopen.
- Discharging an inpatient excludes future recurrence and current active listings after effective date; finalized past events unchanged; no attendance-as-absence fabricated.
- New therapist/student appears in appropriate choice lists; retiring assigned provider is blocked or requires explicit conflict/reassignment resolution; historical IDs remain stable.
- Student placement-end transition handles all affected recurring bookings and leaves no orphaned active assignments.
- A complete synthetic sequence: register -> schedule -> conflict -> absence/replacement -> return -> end/discharge -> save/reopen -> history/statistics. Confirm operational screens agree.
- Must pass Python tests, isolated native Excel smoke, and independent read-only audit. No modification of frozen CC or real workbook and no 2027 patient roster import while gate is closed.

**Status:** requirements frozen; implementation and native tests pending. This documentation is not evidence of functionality.
