# Patient-centric slot-group audit

The timetable source of truth remains one independent assignment per patient.

A visual slot group is derived when assignments share the same therapist and clock time. Day patterns remain attached to each patient assignment.

Rules:

- Same therapist + same time + non-overlapping day patterns can render as a clean shared slot/pair.
- Same therapist + same time + overlapping weekdays are surfaced for review. They are not silently treated as a clean pair.
- Changing one patient's therapist, time, or day pattern rebuilds the visual groups automatically without changing the other patient.
- Daily absence/replacement logic acts on the dated patient occurrence, not on the entire visual cell.

Run against the real workbook with:

```powershell
python Python/tools/audit_slot_groups.py
```
