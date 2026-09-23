# Psychology source schema

This change adds psychology as a source-level treatment specialty without changing the final visual timetable layout.

## SETTINGS

Existing columns 1-8 remain unchanged. A new optional column is appended:

9. `PSYCHOLOGISTS`

The Python reader remains backward-compatible with workbooks that still have only the original eight SETTINGS columns.

## PATIENT_PLANNER

No existing source column is shifted. The psychology block is appended after the current last column:

- `Ψυχ_Ώρα`
- `Ψυχ_Ημέρες`
- `Ψυχ_Ψυχολόγος`

A psychology entry is imported as treatment `Ψυχ` with the named psychologist as its provider.

Because the cross-specialty engine is patient-centric, a psychology commitment blocks the same patient from being scheduled at the same date/time for another specialty.

## Preview rule

`Python/tools/preview_psychology_schema.py` creates `Excel/previews/PSYCHOLOGY_SCHEMA_PREVIEW.xlsm` from the baseline workbook using Microsoft Excel/COM. It never edits the baseline workbook and reports whether source bytes and VBA were preserved.

This feature does not redesign or modify the final operational timetable layout.
