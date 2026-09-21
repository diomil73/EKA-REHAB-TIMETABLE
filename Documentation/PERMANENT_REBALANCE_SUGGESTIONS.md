# Permanent rebalance suggestions

The permanent rebalance planner is read-only. It searches for recurring patient
moves that satisfy two gates simultaneously:

1. Moving the assignment must actually reduce an over-capacity weekday on the
   source provider.
2. The proposed destination provider/time must pass weekly capacity and
   same-day conflict validation for every weekday in the patient's recurring
   day-pattern.

Physiotherapists use the confirmed hard limit of 6 occupied timeslots per day.
Student capacity remains 5 timeslots per day and will be included when the
real student registry is connected to the permanent-assignment workflow.

Pairs remain derived views. If two different patients share one clock time on
complementary weekdays, moving only one patient frees only that patient's
weekdays, not the whole weekly slot.

This planner never changes PATIENT_PLANNER and never selects a patient on the
user's behalf.
