# v27.1 read-only import audit

Date: 2026-09-17

This audit was produced by the Python read-only Excel adapter. The workbook was opened with `data_only=True` and was **not saved or modified**.

## Confirmed import source

- `PATIENTS`: patient registry and patient state.
- `PATIENT_PLANNER`: recurring base programme.
- `SETTINGS`: timeslot grid, day combinations, statuses and existing configuration values.
- `SESSIONS`: audited only. It is **not** used as the Python scheduling source at this stage.

## Baseline findings

- Active patients in `PATIENTS`: **90**.
- `PATIENTS` and the calculated patient identity columns in `PATIENT_PLANNER`: **0 PatientID/name mismatches**.
- Standard timeslots in `SETTINGS`: **7** (`08:30`, `09:15`, `10:00`, `10:45`, `11:30`, `12:15`, `13:00`).
- Existing patient-status values in `SETTINGS`: **8** and are imported without renaming.
- Day combinations in `SETTINGS`: **31**, including `Καθ/να`.
- Timed programme entries with a usable day pattern: **144**.
- Provider-owned (`ΦΘ` / `Ρομποτικό`) entries with a usable day pattern: **85**.
- Timed `PATIENT_PLANNER` entries with no day pattern: **28**. Python does not guess their meaning; they remain an audit item until their rule is confirmed.
- `SESSIONS` contains **77 PatientID/name identity drifts** relative to `PATIENTS`. For that reason, it is not used as the current Python import source.

## Student note

`SETTINGS!THERAPISTS_FTH` still contains the legacy value `φοιτ 1`. The adapter reads existing settings literally, but this value is not treated as the final student registry. The new Python student model remains separate from the therapist role.

## Safety rule

The read-only adapter must never save the `.xlsm`. Excel remains the owner of VBA, formatting, drawings, validations and workbook UI. Any future write integration will use a separate controlled adapter and will be regression-tested before use on the working workbook.
