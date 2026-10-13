# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V71.xlsm`
- V71 builds and opens after the earlier installer hardening.
- Next native preview to build: `MASTER_TOOLBAR_SMOKE_V73.xlsm` from clean V48.

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- The former `Subscript out of range` during MASTER row lookup is fixed.
- The ΦΘ self-conflict false alarm is fixed in native use.
- V71 conflict behavior appears functionally correct in native use: overlap is detected and Save presents the intended conflict decision flow.

## Conflict-validation behavior

A common `validate-vba` bridge command checks pairwise patient-program overlaps using the duration contract:

- standard treatments: 45 minutes,
- Ανακλινόμενο: 30 minutes,
- only shared weekdays count,
- adjacent intervals are allowed.

The patient-program UI:

- validates before **Αυτόματη συμπλήρωση**,
- validates before any atomic Save write begins,
- shows **Συνέχεια** and **Επιστροφή** on Save conflicts,
- **Επιστροφή** cancels before writes,
- **Συνέχεια** is an explicit override.

## V71/V72 presentation fixes

### Conflict dialog runtime layout

The Save conflict dialog was too short and clipped the message/buttons.

Implemented:

- `patient_program_conflict_layout_vba.py`,
- runtime form sizing in `Configure()`,
- explicit label/button placement independent of VBE designer geometry persistence.

### Existing MASTER decimal times

The Ρομποτικό time for patient Γιαννιά appeared as a raw Excel decimal.

Implemented:

- preview installer iterates all nonblank PatientIDs in `PATIENTS`,
- invokes `RefreshMasterPatientProgram` for each existing patient,
- re-renders all treatment cells through the composite `hh:mm` renderer.

Focused gate for these fixes:

`57 passed in 3.96s`

## New native scheduling observations

Two further issues were reported while exercising the patient-program flow.

### 1. Manual ΦΘ provider/time conflict message was too generic

Desired behavior: if the selected therapist is already occupied, identify the concrete blocker, e.g. weekday, time, therapist and blocking patient/treatment.

Implemented in source:

- manual ΦΘ failure diagnostics inspect provider occupancy,
- when a concrete patient booking explains the failure, the reason includes weekday, time, therapist, patient name and treatment,
- generic fallback remains only when no concrete blocker can be resolved.

### 2. False Ανακλινόμενο conflict during multi-treatment edit

Observed scenario:

- ΦΘ daily at `11:30`, duration 45',
- Ανακλινόμενο `Δε-Τε-Πα` at `11:00`, duration 30',
- these are exactly adjacent (`11:00-11:30` and `11:30-12:15`) and must be legal.

Root cause:

- during automatic completion of an edit draft, the solver compared a new treatment against stale persisted assignments of other treatments from the same patient,
- it ignored only the old assignment of the same treatment,
- therefore the new Ανακλινόμενο could collide with the patient's old persisted ΦΘ even though the draft's new ΦΘ was at 11:30.

Fix implemented in `automatic_patient_schedule.py`:

- all old persisted assignments whose treatment participates in the current draft are removed from the existing-patient conflict baseline,
- the new draft assignments are then checked jointly against each other,
- exact adjacency remains allowed,
- true joint conflicts return detailed weekday/treatment/time ranges.

Regression coverage includes:

- stale old ΦΘ no longer blocking new `Ανακλινόμενο 11:00` + `ΦΘ 11:30`,
- manual occupied ΦΘ slot naming day/provider/patient,
- detailed joint conflict ranges.

## Latest confirmed test checkpoint

User-run gates after the latest scheduling changes:

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`

The wider gate covered:

- `Tests/test_automatic_patient_schedule.py`
- `Tests/test_patient_program_bridge.py`
- `Tests/test_patient_program_form_vba.py`
- `Tests/test_patient_program_master_render_vba.py`
- `Tests/test_patient_registration_scheduling_slice.py`

Source/integration verification is therefore green. Native V73 verification is next.

## Immediate next start point

1. Build `MASTER_TOOLBAR_SMOKE_V73.xlsm` from clean V48.
2. Native re-test:
   - manual ΦΘ occupied therapist slot gives precise patient/day/time diagnostic,
   - `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15` is accepted,
   - true treatment overlaps identify exact therapies/day/time,
   - conflict dialog shows full message and both buttons,
   - Γιαννιά Ρομποτικό time appears as `hh:mm`.
3. Then continue with edit-existing-program and optional manual names.
4. Full regression suite only after native edit smoke is clean.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual day/time/provider constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration currently used by the solver: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal and is not an overlap.
- Recliner adjacency to ΦΘ is a strong soft preference, not a hard rule.
- MASTER is a derived view; `PATIENT_PLANNER` remains authoritative for the inpatient patient program.
- Conflict override must be explicit. Returning for correction must perform no writes.

## Recovery order if context is lost

Read these in order:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. tests around the currently active patient-program modules

Use `python -m pytest` on Windows, never bare `pytest`.
