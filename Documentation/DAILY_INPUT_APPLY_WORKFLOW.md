# DAILY_INPUT → Operational timetable

This workflow is the authoritative daily-input business path.

## Locked production rule

`DAILY_INPUT` must not close and reopen the workbook in normal use.

Production interaction is now:

1. enter absences/cancellations in `DAILY_INPUT`;
2. request application;
3. save and validate the current workbook state;
4. compute the dated effective state including existing ACTIVE replacement facts;
5. build one operational write plan containing the therapist/day presentation and the MASTER daily overlay;
6. serialize only validated cell patches;
7. apply those patches to the already-open workbook;
8. read back every value/presentation patch;
9. refresh and verify the therapist workload strip from the updated `THERAPIST_DAILY` projection;
10. if therapist-absence sessions remain unresolved, run the global replacement solver in place;
11. persist ordinary replacement facts automatically; show guided double-booking review only when an explicit operator decision is genuinely required;
12. save the same workbook and remain in the application shell.

`Python/tools/prepare_daily_input_in_place.py` is the read-only preparation layer. It rebuilds the effective daily state, includes ACTIVE `DAILY_REPLACEMENT_FACTS`, builds the patient/therapist presentation plus `MASTER_SCHEDULE` overlay, and emits the validated `WritePlan` as a UTF-8 payload. It does not create or replace an authoritative workbook copy.

`Python/rehab_excel/daily_input_action_vba.py` applies the payload in-session. It mirrors the native presentation semantics for values, clears, fills, borders, font roles, wrap text, minimum font size and rich-text runs. Success is not reported until the patch read-back completes and the operational workload view has refreshed.

Before mutation, VBA creates a temporary `SaveCopyAs` recovery file. This copy is rollback/backup only. The live workbook is never closed or replaced. If any patch or read-back fails, affected cells are restored from the recovery copy and the action reports failure.

## Therapist absence / replacement continuation

A therapist absence must never route the operator to a raw `REPLACEMENTS` worksheet.

The production continuation is:

`DAILY_INPUT -> effective state -> global replacement solver -> ordinary assignments -> guided double-booking review only if required -> remaining treatment suspension -> final projections -> workload refresh`

`Python/tools/prepare_global_replacements_in_place.py` is read-only. It reuses the authoritative global solver, provider policy, service hours, patient cross-specialty availability, existing replacement facts, student eligibility and reclining adjacency rules.

It emits:

- validated `REPLACEMENT` facts for ordinary assignments;
- `SUSPENSION` records only when no reviewable exception exists;
- final `PATCH` rows for `THERAPIST_DAILY` / patient-centred views and `MASTER_SCHEDULE` when the state is fully resolvable;
- `review_required=true` instead of silently suspending a session when an explicit double-booking exception is available.

`Python/rehab_excel/global_replacement_action_vba.py` persists this result in the already-open workbook with duplicate checks, read-back, rollback protection and hidden diagnostics. If review is required it opens the existing guided `frmDoubleBookingReview`; it does not expose a raw worksheet.

## Operational projection success gate

The action is not considered operationally complete merely because `DAILY_INPUT` was parsed successfully.

The following must be current after application:

- `THERAPIST_DAILY` reflects the dated effective session state;
- `MASTER_SCHEDULE` reflects the same dated state through the MASTER daily overlay;
- therapist workload badges are rebuilt from `THERAPIST_DAILY` and continue to count **occupied timeslots**, not patient lines;
- already-active replacements remain part of the effective state and are not silently lost on re-apply.

MASTER presentation rules are now explicit:

- `REPLACED` shows the effective time/provider;
- `PATIENT_ABSENT` shows `ΑΠΩΝ` plus the reason when available;
- `CANCELLED` shows `ΑΚΥΡΩΘΗΚΕ` plus the reason when available;
- unresolved `THERAPIST_ABSENT` shows `ΑΝΑΜΟΝΗ ΑΝΑΠΛΗΡΩΣΗΣ` until the replacement continuation resolves it;
- when all dated sessions for an inpatient are non-active, the MASTER patient status is overlaid as non-active as well.

Workload counting rules are also explicit:

- a non-empty cell is not automatically occupied;
- struck-through absence/cancellation lines do not consume a slot;
- italic replacement-source lines do not consume the original therapist slot;
- `→` routing annotation lines do not consume a slot;
- a shared slot with at least one active patient still counts as exactly one occupied timeslot.

The V4 native smoke exposed the raw replacement-sheet routing gap. The V5 smoke exposed a second integration gap: a real cancellation (`ΓΙΑΝΝΙΑ`, reason `Διακομιστήριο`) left MASTER visually active and kept `ΠΑΠΑΜΑΝΟΥ` at `6/6`. Root cause was an overlay that handled only replacements plus a workload counter that treated every non-empty cell as occupied. Both are now regression-guarded.

## Diagnostic-first native testing

Every native production action must have hidden diagnostics before its first smoke test. Operators see only clear instructions/results; technical details go to temp logs for troubleshooting.

Current logs include:

- `%TEMP%\eka_daily_input_last_error.log`
- `%TEMP%\eka_replacement_last_error.log`

If a native smoke fails, read the relevant diagnostic before another retry. Do not perform repeated blind tests.

The V3/V4 debugging sequence confirmed why this rule matters: a `SaveCopyAs` failure was traced immediately to a doubled runtime backslash in a TEMP path. Runtime VBA strings are now guarded against accidental doubled backslashes produced by Python raw-string source.

## Development/smoke previews

Developer preview tools may still create disposable copies. Those are test/install/migration tools only and do not define normal operator behavior.

A fresh application preview must explicitly reinstall and verify the operational modules used by DAILY_INPUT continuation, including global replacement action, guided double-booking review, and therapist workload strip. Pre-cleaning a component without reinstalling it is a build failure, not an acceptable fixture dependency.

## Business rules

- Selecting a patient/therapist name activates the row.
- `Όλη ημέρα = ΝΑΙ` ignores time fields.
- A patient timeslot affects only that dated occurrence.
- A therapist row with only `Από` means exactly that timeslot. Users do not need artificial ranges such as `08:30-08:31`.
- Patient status is the absence/cancellation reason. `ΠΑΡΩΝ` does not create an absence.
- Pair/group cells are rendered patient-centrically, so only the affected patient line is changed.
- Patient absence wins over therapist absence for the same session.
- The global solver handles all surviving therapist-absence sessions together instead of greedily one-by-one.
- Existing ACTIVE replacement facts are preserved.
- Double booking is explicit-only: the operator must accept a presented exception before it becomes an active replacement.
- The base schedule remains the source of truth.
- Daily facts remain overlays and must not destructively rewrite recurring/base scheduling.
