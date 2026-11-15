# DAILY_INPUT → Operational timetable

This workflow is the authoritative daily-input business path.

## Locked production rule

`DAILY_INPUT` must not close and reopen the workbook in normal use.

Production interaction:

1. enter absences/cancellations in `DAILY_INPUT`;
2. apply the daily change;
3. validate the current workbook state;
4. compute the dated effective state including ACTIVE replacement facts;
5. build the therapist/day and MASTER daily projections;
6. apply validated patches to the already-open workbook;
7. read back the written values/presentation;
8. refresh workload and guided replacement continuation when needed;
9. save the same workbook and remain in the application shell.

Applying a DAILY_INPUT change is a committed operator action and therefore saves automatically. It does not depend on a separate manual Save button.

## MASTER presentation

`MASTER_SCHEDULE` remains the recurring/week view. Daily overlays must not destroy its recurring treatment text.

Locked presentation:

- `REPLACED` may show the dated effective time/provider;
- patient absence/cancellation keeps the recurring treatment text and applies a **red strike-through** to treatment cells that actually occur on the selected date;
- the final `Κατάσταση` column carries the dated patient reason/status, e.g. `Διακομιστήριο`;
- an all-day absence affects every materialized session for that patient on the selected date, but it must not visually cancel recurring treatments belonging only to other dates of the week;
- pending therapist replacement is an operational daily overlay, never a destructive recurring edit.

The V6 native smoke proved that replacing a therapy cell with the text `ΑΠΩΝ` was the wrong presentation. The treatment cell must remain recognizable and be red-struck instead.

## Therapist/day projection and workload

`THERAPIST_DAILY` reflects the same dated effective state. Workload remains based on **occupied timeslots**, not patient lines.

Rules:

- struck-through patient absence/cancellation lines do not consume a therapist slot;
- italic replacement-source lines do not consume the original therapist slot;
- `→` routing annotation lines do not consume a slot;
- a shared slot with at least one active patient counts as one occupied timeslot;
- workload is attributed from the authoritative original/effective therapist mapping, never inferred from a screenshot.

The V6 observation that `ΠΑΠΑΜΑΝΟΥ` remained `6/6` is not by itself treated as a confirmed defect because the user immediately noted that the therapist mapping may have been mistaken. Future native verification must check the actual therapist/session mapping before judging the badge.

## Navigation and controls

The application shell must preserve usable navigation:

- `Φόρτος θεραπευτών` on MASTER opens `THERAPIST_DAILY`;
- `THERAPIST_DAILY` has a fixed `Επιστροφή στο MASTER` control;
- workload button/badges are locked shapes and MASTER drawing objects are protected so operators cannot drag the workload row accidentally;
- the therapist sheet uses its two existing layout blocks; do not create artificial horizontal scrolling or expose raw sheets as a workaround.

## Reversible DAILY_INPUT changes

Every committed DAILY_INPUT absence/cancellation must be traceable and reversible.

Each active source row now has a stable daily-change identity containing:

- date;
- source row;
- change type (`patient_absence`, `session_cancellation`, `therapist_absence`);
- subject/display name;
- reason;
- all-day/timeslot semantics.

These changes are represented in the existing `DAILY_CHANGE_LEDGER` UI alongside replacement changes. DAILY_INPUT rows use dedicated metadata and their individual `Επαναφορά` button routes to the DAILY_INPUT restore transaction, not the replacement restore backend.

Restore semantics:

1. identify the exact DAILY_INPUT source row from ledger metadata;
2. prepare a read-only full-day projection as if only that source row were removed;
3. rebuild the complete `THERAPIST_DAILY` grid from recurring/base truth plus surviving daily overlays;
4. reset MASTER daily presentation and reapply only surviving overlays;
5. clear the exact DAILY_INPUT source row in the open workbook;
6. apply/read-back the rebuild;
7. mark the ledger change `RESTORED`;
8. refresh workload/navigation and save;
9. on any failure restore from the rollback copy.

This remains an in-session production flow. There is no close/reopen or live-file replacement.

Current restore diagnostic:

- `%TEMP%\eka_daily_input_restore_last_error.log`

## Therapist absence / replacements

A therapist absence must never route the operator to a raw `REPLACEMENTS` worksheet.

Production continuation:

`DAILY_INPUT -> effective state -> global replacement solver -> ordinary assignments -> guided double-booking review only if required -> remaining treatment suspension -> final projections -> workload refresh`

Existing ACTIVE replacement facts remain part of the effective state. Double booking remains explicit-only.

## Diagnostic-first native testing

Every native production action must have hidden diagnostics before its first smoke test. After a generic operator failure, read the relevant log before any retry.

Current logs include:

- `%TEMP%\eka_daily_input_last_error.log`
- `%TEMP%\eka_replacement_last_error.log`
- `%TEMP%\eka_daily_input_restore_last_error.log`

Do not repeat blind native tests.

## Development/smoke previews

Developer preview tools may create disposable copies. Normal operators do not use preview/replace workflows.

Fresh previews must deterministically reinstall and verify the complete operational set: shell, MASTER toolbar, DAILY_INPUT action, global replacements, guided double-booking review, workload/navigation, in-session daily-change ledger and reversible DAILY_INPUT ledger extension.

## Business rules

- Selecting a patient/therapist name activates the row.
- `Όλη ημέρα = ΝΑΙ` ignores time fields.
- All-day patient absence covers every scheduled/materialized session for that date.
- A patient timeslot absence affects only that dated occurrence.
- A therapist row with only `Από` means exactly that timeslot.
- `ΠΑΡΩΝ` does not create an absence.
- Patient absence wins over therapist absence for the same session.
- Pair/group cells remain patient-centric.
- The global solver handles surviving therapist-absence sessions together.
- Base/recurring schedule remains the source of truth.
- Daily facts are overlays and are never destructive recurring edits.
