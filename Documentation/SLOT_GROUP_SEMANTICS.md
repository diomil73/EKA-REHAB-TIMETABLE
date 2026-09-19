# Slot-group semantics

The timetable remains patient-centric. Each recurring assignment belongs to one patient.

A **derived pair** requires:

- the same therapist,
- the same clock time,
- exactly two **different** patients,
- no overlapping weekdays.

Two complementary assignments for the **same patient** are not a pair. They are a
**same-patient split schedule** and remain separate source assignments. This matters when
different day patterns carry different operational attributes, such as robotic vs standard
physiotherapy.

Overlapping-day groups remain review items and are never silently converted into clean pairs.
