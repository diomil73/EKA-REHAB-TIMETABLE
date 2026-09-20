# Capacity rules

Confirmed operational limits:

- Physiotherapist: maximum **6 distinct timeslots per day**.
- Student: maximum **5 distinct timeslots per day**.

Capacity is counted by distinct clock time, not by number of patient rows. If two patient rows legitimately share one clock slot, that consumes one capacity slot for that provider on that day.

The rule applies to:

- daily replacements,
- new patient placement,
- permanent therapist changes,
- permanent time changes,
- future schedule editing workflows.

For recurring assignments, capacity is checked separately for every weekday in the proposed day pattern. A new assignment is rejected if even one weekday would exceed the provider limit.

Capacity validation is separate from conflict validation. An existing patient at the same provider/time/day may not increase the distinct-timeslot count, but it is still a scheduling conflict and must be handled by the conflict engine.
