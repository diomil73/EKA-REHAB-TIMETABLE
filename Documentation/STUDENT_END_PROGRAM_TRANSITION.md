# Student end-program transition

This stage adds a read-only planning layer for student placement completion.

Rules implemented:

- while a student is active, the real display name remains visible;
- after `placement_end`, the timetable display label becomes `Φοιτ.N` while the real `student_id` and identity remain unchanged internally;
- student operational assignments can be resolved back to their recurring base entries;
- the planner reports the affected recurring patients/entries;
- permanent takeover suggestions are generated only when they pass the existing weekly provider-capacity, provider-collision, and cross-specialty patient-conflict checks;
- the current supervising therapist may appear as a valid takeover option when the recurring slot is safe;
- no assignment is changed automatically and no workbook write is performed by this module.

The next integration step is to connect this planner to the workbook/manager workflow so the user can review and explicitly accept a permanent reassignment.
