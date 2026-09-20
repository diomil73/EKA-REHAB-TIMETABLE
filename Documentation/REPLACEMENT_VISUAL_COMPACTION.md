# Replacement visual compaction

`THERAPIST_DAILY` is already a view of one specific operational day. Therefore a
replacement note in the **original** therapist cell does not repeat the date.

Source cell:

```text
ΗΛΙΑ ΧΡΥΣΟΥΛΑ
→ Καλυβιώτης 09:15
```

The original patient line remains italic/muted for a replacement, not struck
through. Strikethrough remains reserved for an absence/cancellation where the
session will not take place.

The **destination** cell still carries the temporary date marker because it is
inserted into a recurring timetable cell and must not look like a permanent base
assignment:

```text
ΗΛΙΑ ΧΡΥΣΟΥΛΑ [Πα 18/09]
```

For a paired/grouped source cell, only the affected patient receives the
replacement note. Other group members stay unchanged.
