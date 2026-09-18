# Excel write-back safety

## Current stage: dry run only

Python can describe and validate intended workbook changes, but it **does not
save or modify the `.xlsm`**.

The dry-run layer now applies two independent gates:

1. **sheet source contract**: the sheet must be an approved future output/view;
2. **cell layout contract**: the exact target cell must be inside a confirmed
   write zone.

The workbook byte hash is checked before and after the dry run and must remain
unchanged.

## Protected authoritative / legacy sheets

Python write-back remains blocked for:

- `PATIENTS`
- `PATIENT_PLANNER`
- `SETTINGS`
- `REPLACEMENT_LOG`
- `THERAPIST_ATTENDANCE`
- `DAILY_LOG`
- `SESSIONS`
- `NEW_PATIENT`

## Confirmed cell-level dry-run zones

- `THERAPIST_DAILY!B2:J8`
- `THERAPIST_DAILY!B12:J18`
- `MASTER_SCHEDULE!D2:J5000`
- `CONFLICT_LOG!A2:A5000`
- `REPLACEMENTS!A1:V5000`

`STATISTICS` is still classified as a future derived output at sheet level,
but has **no** confirmed cell-level write zone yet and is therefore blocked.

## Protected layout inside writable sheets

Examples:

- `THERAPIST_DAILY` therapist headers, time column and refresh button area are
  protected;
- `MASTER_SCHEDULE` patient identity/room/status columns are protected;
- `CONFLICT_LOG!A1` title is protected.

## Why no direct workbook save yet

The production workbook is macro-enabled and contains VBA, form controls,
DrawingML shapes and existing presentation rules. The project does not enable
a generic package save path for production write-back. Real mutation will be
introduced only through an Excel-hosted automation layer after the render
contract and formatting roles are tested on a disposable working copy.

## Next gate

Before enabling a real write:

- map `DailySessionState` to the confirmed `THERAPIST_DAILY` cells;
- define presentation roles for normal, absent, replacement, infectious,
  robotic and active-student states;
- generate a complete write plan for one test day;
- apply it only to a disposable working copy through Excel-host automation;
- verify formatting, buttons, VBA and workbook integrity in desktop Excel.
