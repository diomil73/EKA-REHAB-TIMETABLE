# Patient-centric timetable contract

## Core decision

The source of truth is **one independent recurring assignment per patient**.
A pair, trio, or any other shared timetable cell is **derived presentation**, not source data.

A recurring patient assignment owns at least:

- PatientID
- therapist/provider
- start time
- day pattern
- treatment
- special attributes such as robotic

The current `BaseScheduleEntry` already represents this model closely enough to remain the base contract.

## Automatic visual grouping

For therapist timetable rendering:

> same therapist + same time -> same derived visual SlotGroup

Each member keeps its own day pattern. Examples:

- Patient A -> therapist X -> 12:15 -> `Δε-Τε-Πα`
- Patient B -> therapist X -> 12:15 -> `Τρ-Πε`

They form one conflict-free visual pair automatically.

The pair itself is never edited or stored as an independent scheduling object.

## Changing one member

If Patient A changes therapist, time, or days, only Patient A's assignment changes. The timetable groups are then rebuilt.

Therefore:

- old pairs dissolve automatically;
- new pairs form automatically;
- no special "break pair" operation is required;
- one patient's move cannot silently rewrite the other patient's assignment.

## Daily operations inside a shared cell

A shared visual cell does not mean the members share the same dated session.
The day pattern determines which patient has treatment on a specific date.

Consequences:

1. If today's patient is absent, only that patient's rendered line/occurrence is marked absent.
2. If the therapist is absent today, only the patient(s) whose day pattern applies today need a daily replacement.
3. A longer therapist absence can generate affected occurrences across a date range, still patient by patient.
4. If all members of a visual group are affected across their respective dates, they may end up moving together, but this is the result of independent assignments, not a group-level source record.

## Conflict behaviour

Assignments with the same therapist and time but overlapping weekdays are kept visible in the same derived slot group and are marked as overlapping. They are not silently accepted as a valid pair.

This lets the conflict engine distinguish:

- complementary members that safely share a visual slot;
- actual same-day collisions that require attention.
