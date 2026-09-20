# Capacity rebalance

Confirmed daily capacity rules:

- Physiotherapist: maximum **6 distinct active timeslots per day**.
- Student: maximum **5 distinct active timeslots per day**.
- Multiple patients sharing the same clock time count as one occupied timeslot.
- A timed base assignment with blank day-pattern is treated as `Καθ/να` (Monday-Friday).

## Existing over-capacity programme

The system must **not automatically move an existing patient** merely because the imported baseline is already over capacity. Instead it should:

1. flag the provider as over capacity,
2. block any additional assignment that would worsen the excess,
3. show the active slots that could be rebalanced,
4. show feasible replacement options for each patient,
5. leave the final choice to the user.

If two active patients occupy the same provider/time on the same date, moving only one does **not** free that timeslot. Every active session in that slot must be moved before the provider's occupied-timeslot count falls.
