# Empty day-pattern means daily

Confirmed operational rule:

- If a schedule assignment has a valid time but its day-pattern cell is empty, it means **Καθ/να**.
- `Καθ/να` means Monday through Friday (`Δε-Τρ-Τε-Πε-Πα`).
- This is a data rule, not a display guess. The Excel reader normalizes such entries to `Καθ/να` before scheduling, workload, replacement ranking or capacity validation.
- Unknown **non-empty** day-patterns still stop validation instead of being guessed.

## Capacity consequence

These implicit daily assignments count toward every weekday's provider capacity:

- physiotherapist: maximum **6 distinct timeslots/day**
- student: maximum **5 distinct timeslots/day**

Two patients sharing the same clock time on complementary days are independent assignments, but on any single date only the active occurrence counts. Multiple patients genuinely sharing one clock time on the same day count as multiple sessions but one occupied timeslot for the capacity ceiling; conflict rules remain separate.
