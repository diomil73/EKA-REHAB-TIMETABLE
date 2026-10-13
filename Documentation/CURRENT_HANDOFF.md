# Current Handoff

Last updated: 2026-10-13

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base for patient-scheduling builds: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested successfully by the user: `MASTER_TOOLBAR_SMOKE_V69.xlsm`

## What is verified in native Excel

- The new patient-centric multi-treatment form opens and is usable.
- Legacy single-treatment VBA no longer blocks the preview with the earlier syntax error.
- First patient-program save succeeds.
- MASTER patient-program rendering works after save instead of showing raw decimals or literal `hh:mm`.
- MASTER rendering uses composite text from `PATIENT_PLANNER`: time + days + provider/name where applicable.
- The former `Subscript out of range` during MASTER row lookup is fixed by guarding empty MASTER patient cells before `Split`.
- Final save can succeed after changing a treatment time.

## Current small regression under verification

Scenario: change the time of a non-ΦΘ specialty in an existing program and press **Αυτόματη συμπλήρωση**. V69 reported a false ΦΘ problem even though ΦΘ had not been changed; normal Save still succeeded.

The first fix correctly found the existing same-patient/same-treatment `base_entry_id` and passed it as `ignore_session_id`. The focused test exposed that the core candidate engine only ignored that session for patient availability, not therapist availability/workload.

Deep fix now implemented:

- `rank_therapist_candidates()` builds `effective_sessions` excluding `ignore_session_id`.
- The same effective set is used for therapist workload, therapist availability, and patient availability.
- The session being replaced therefore does not occupy its own therapist slot or capacity.
- Real sessions from other patients remain blocking.

Core regression test now checks both slot release and workload release for the ignored session.

## Latest test checkpoint

Latest user-reported focused run:

`24 passed, 1 failed in 3.73s`

The single failure was `test_existing_same_patient_treatment_is_ignored_during_edit_revalidation`, showing the therapist side still counted the old session. The deeper core fix was committed after that run and is **not yet test-verified**.

## Immediate next start point

1. Pull latest `feature/outpatient-scheduling`.
2. Run:

   `python -m pytest Tests/test_scheduling_candidates.py Tests/test_workbook_scheduling_candidates.py Tests/test_patient_recurring_scheduling.py Tests/test_automatic_patient_schedule.py -q`

3. If green, build fresh preview from V48, using the next preview number after V69.
4. Re-test:
   - existing patient program,
   - change one non-ΦΘ treatment time,
   - press **Αυτόματη συμπλήρωση**,
   - unchanged ΦΘ must no longer be falsely reported.
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
