# Excel write-back safety

## Current stage: dry run only

Python can now describe and validate intended workbook changes, but it **does not save or modify the `.xlsm`**.

The dry-run layer exists to prove three things before real Excel automation is enabled:

1. only approved output/view sheets are targeted;
2. authoritative source sheets stay protected;
3. the workbook byte hash is unchanged after the dry run.

## Protected sheets

Python write-back is currently blocked for authoritative and legacy operational sources, including:

- `PATIENTS`
- `PATIENT_PLANNER`
- `SETTINGS`
- `REPLACEMENT_LOG`
- `THERAPIST_ATTENDANCE`

## Current approved future targets

The source contract currently marks these as future Python output targets:

- `MASTER_SCHEDULE`
- `THERAPIST_DAILY`
- `REPLACEMENTS`
- `CONFLICT_LOG`
- `STATISTICS`

A target must also physically exist in the workbook before a write plan is accepted.

## Why no openpyxl save

The production workbook is macro-enabled and contains VBA/UI structures. The project does not use `openpyxl.save()` as the future production write path. Real write-back will be introduced later through Excel-controlled automation after the render contract is fixed and tested.

## Next gate

Before enabling a real write:

- map Python operational states to exact Excel cells/ranges;
- define formatting roles (infectious yellow, robotic pink, active student green, strike-through overlay);
- test changes on a disposable working copy;
- verify workbook integrity and VBA behavior in desktop Excel.
