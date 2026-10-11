# Current Handoff

Last updated: 2026-10-11

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested successfully by the user: `MASTER_TOOLBAR_SMOKE_V69.xlsm`

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- MASTER patient-program rendering works after save instead of showing raw decimals or literal `hh:mm`.
- MASTER rendering now uses composite text from `PATIENT_PLANNER`: time + days + provider/name where applicable.
- The former `Subscript out of range` during MASTER row lookup is fixed by guarding empty MASTER patient cells before `Split`.
- Final save can succeed even after changing a treatment time.

## Latest small issue observed

After changing the time of a non-ΦΘ specialty and pressing **Αυτόματη συμπλήρωση**, the UI reported a false ΦΘ problem even though ΦΘ had not been changed. Pressing normal Save afterwards succeeded.

This was diagnosed as therapist-candidate self-conflict during edit/revalidation: the already stored ΦΘ session for the same patient/treatment was being counted as occupied against itself.

Code fix implemented:

- identify the existing `base_entry_id` for the same `PatientID` + treatment,
- pass it as `ignore_session_id` into the therapist candidate engine,
- preserve real conflicts with other patients/sessions.

This fix is **implemented but not yet re-verified in native Excel**.

## Latest confirmed test checkpoint

Latest user-reported focused green gate before the final self-conflict fix:

`44 passed in 3.84s`

Do not claim the newest self-conflict fix is test-verified until the focused suite is rerun.

## Immediate next start point

1. Pull latest `feature/outpatient-scheduling`.
2. Run focused tests for:
   - `Tests/test_workbook_scheduling_candidates.py`
   - `Tests/test_patient_recurring_scheduling.py`
   - `Tests/test_automatic_patient_schedule.py`
   - relevant patient-program tests if needed.
3. Build a fresh preview from V48, using the next preview number after V69.
4. Re-test only the small scenario:
   - existing patient program,
   - change one non-ΦΘ treatment time,
   - press **Αυτόματη συμπλήρωση**,
   - confirm unchanged ΦΘ is no longer falsely reported.
5. Then move to the still-unverified edit workflow:
   - **Βασικό μενού → Επεξεργασία ασθενή**,
   - save registry changes,
   - same patient-program form opens,
   - existing program is preloaded,
   - modify and atomic-save it.
6. Verify optional manual names for Ρομποτικό / Εργο / Λογο / Ψυχ persist and reload.
7. Verify duration rule ΦΘ 45' + Ανακλινόμενο 30', especially ΦΘ 09:15 and recliner preference for 10:00 when available.
8. Run the full regression suite only after the native edit smoke is clean.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual day/time/provider constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration currently used by the solver: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Recliner adjacency to ΦΘ is a strong soft preference, not a hard rule.
- MASTER is a derived view; `PATIENT_PLANNER` remains authoritative for the inpatient patient program.

## Recovery order if context is lost

Read these in order:

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. tests around the currently active patient-program modules

Use `python -m pytest` on Windows, never bare `pytest`.
