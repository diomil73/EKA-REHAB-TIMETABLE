# Native Excel write-back contract

## Purpose

Python may now prepare an actual `THERAPIST_DAILY` preview **only in a new `.xlsm` copy**. The baseline/source workbook must never be the destination.

## Why native Excel

The project does not save production `.xlsm` files with `openpyxl`. Microsoft Excel itself performs the save through Windows COM so VBA, drawings, buttons and workbook-native structures remain under Excel's control.

## Safety rules

1. The source and output paths must be different.
2. Only verified sheet/cell write zones are accepted.
3. `PATIENTS`, `PATIENT_PLANNER`, `SETTINGS` and the current legacy operational sheets remain protected.
4. The source SHA-256 is checked before and after every copy-write.
5. If Excel fails during the operation, the partial preview copy is deleted.
6. The current native writer targets `THERAPIST_DAILY` render output only. `MASTER_SCHEDULE` remains untouched by daily overlays.

## Rich text

A daily cell can carry multiple lines. Character-level formatting is preserved in the write plan so an original assignment can be struck through while a replacement line remains active.

## Presentation roles

- infectious: existing v27 yellow (`FFFFFF43`)
- robotic: existing v27 robotic salmon/pink (`FFF8CBAD`)
- active student: green font
- muted/original line: gray font
- minimum font size: 10
- wrap text: enabled by render plan

## Runtime

Native mutation requires Windows, Microsoft Excel and `pywin32`.

Environment check:

```powershell
python Python/tools/check_excel_runtime.py
```

No real workbook should be used as an output until the generated preview copy has been inspected manually in Excel.
