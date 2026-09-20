# Patient-line formatting inside shared cells

A THERAPIST_DAILY cell can contain more than one patient line. Presentation must therefore remain patient-centric.

Rules:

- Existing recurring patients keep their own day pattern when a temporary replacement is inserted into the same cell.
- Robotic font styling belongs only to the robotic patient's line. It must not spill onto a newly inserted non-robotic patient.
- A temporary replacement line gets robotic orange only when the replacement session itself is robotic.
- Replacement notes on the original cell remain compact.

Example:

```text
ΜΑΚΡΙΑΔΑΚΗ [Τρ-Πε]        <- robotic orange if that assignment is robotic
ΗΛΙΑ ΧΡΥΣΟΥΛΑ [Πα 18/09]  <- normal font when non-robotic
```
