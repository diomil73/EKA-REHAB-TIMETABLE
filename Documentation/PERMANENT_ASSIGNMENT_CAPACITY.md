# Permanent assignment capacity

Before a new patient is placed, or an existing patient is permanently moved to a new therapist/time, the assignment is validated on **every weekday in its recurring day-pattern**.

Confirmed limits:

- Physiotherapist: maximum **6 distinct occupied timeslots/day**.
- Student: maximum **5 distinct occupied timeslots/day**.
- Blank day-pattern with a valid time is imported as `Καθ/να` (Monday-Friday).
- Patients with complementary day-patterns may share the same therapist/time visual cell.
- Two different patients active on the **same weekday at the same therapist/time** are a conflict, even though they occupy one clock-time.
- When moving an existing assignment, the source assignment is removed from the projection before the destination is evaluated.

The check is read-only. A blocked result must never modify the base timetable.
