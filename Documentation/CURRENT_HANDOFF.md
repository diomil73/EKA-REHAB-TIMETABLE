# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested by the user: `MASTER_TOOLBAR_SMOKE_V71.xlsm`
- V71 builds and opens after the earlier installer hardening.
- V72 source batch exists but native verification is not yet complete.

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

Latest confirmed focused gate for these fixes:

`57 passed in 3.96s`

## New native scheduling observations

Two further issues were reported while exercising the current patient-program flow.

### 1. Manual ΦΘ provider/time conflict message is too generic

When the user changes ΦΘ time to a slot where the selected therapist is already occupied, the system should identify the real blocker instead of only saying that no valid combination exists.

Desired diagnostic example:

`Δευτέρα 11:30: ο θεραπευτής Σαρράς είναι απασχολημένος με τον ασθενή Χ (ΦΘ).`

Implemented in source:

- manual ΦΘ failure diagnostics now inspect provider occupancy,
- when a concrete patient booking explains the failure, the reason includes weekday, time, therapist, patient name and treatment,
- generic fallback remains only when no concrete blocker can be resolved.

### 2. False Ανακλινόμενο conflict during multi-treatment edit

Observed scenario:

- ΦΘ daily at `11:30`, duration 45',
- Ανακλινόμενο `Δε-Τε-Πα` at `11:00`, duration 30',
- these are exactly adjacent (`11:00-11:30` and `11:30-12:15`) and must be legal.

Root cause identified:

- during automatic completion of an edit draft, the solver compared a new treatment against stale persisted assignments of other treatments from the same patient,
- it ignored only the old assignment of the same treatment,
- therefore the new Ανακλινόμενο could collide with the patient's old persisted ΦΘ even though the draft's new ΦΘ was at 11:30.

Fix implemented in `automatic_patient_schedule.py`:

- all old persisted assignments whose treatment participates in the current draft are removed from the existing-patient conflict baseline,
- the new draft assignments are then checked jointly against each other,
- exact adjacency remains allowed,
- true joint conflicts now return detailed weekday/treatment/time ranges instead of a generic conflict sentence.

Regression tests added for:

- stale old ΦΘ no longer blocking new `Ανακλινόμενο 11:00` + `ΦΘ 11:30`,
- manual occupied ΦΘ slot naming day/provider/patient,
- detailed joint conflict ranges.

## Latest confirmed test checkpoint

Latest user-run narrow regression gate for the edit-draft replacement + detailed scheduling diagnostics batch:

`16 passed in 2.87s`

This confirms `Tests/test_automatic_patient_schedule.py` is green after the new changes. The wider patient-program integration gate is still pending and should run before any new native preview is built.

## Immediate next start point

1. Run the wider focused patient-program gate:
   - `Tests/test_automatic_patient_schedule.py`
   - `Tests/test_patient_program_bridge.py`
   - `Tests/test_patient_program_form_vba.py`
   - `Tests/test_patient_program_master_render_vba.py`
   - `Tests/test_patient_registration_scheduling_slice.py`
2. If green, build the next fresh preview from clean V48, expected next version: V73.
3. Native re-test:
   - manual ΦΘ occupied therapist slot gives precise patient/day/time diagnostic,
   - `Ανακλινόμενο 11:00-11:30` + `ΦΘ 11:30-12:15` is accepted,
   - true treatment overlaps identify exact therapies/day/time,
   - conflict dialog shows full message/buttons,
   - Γιαννιά Ρομποτικό time appears as `hh:mm`.
4. Then continue with edit-existing-program and optional manual names.
5. Full regression suite only after native edit smoke is clean.

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
