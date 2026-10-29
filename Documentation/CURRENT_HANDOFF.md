# Current Handoff

Last updated: 2026-10-27

## Read order

Start with `Documentation/ONBOARDING.md`, then `Documentation/IN_PLACE_MUTATION_ARCHITECTURE.md`, then `Documentation/DAILY_INPUT_APPLY_WORKFLOW.md`. For the current native checkpoint read `Documentation/V15_THERAPIST_ABSENCE_REMEDIATION.md`.

This file is the exact resume point.

## Branch / trusted state

- Branch: `feature/outpatient-scheduling`
- Baseline `Rehab_Center_System_v27_1.xlsm` is immutable.
- Excel is the application host/UI; Python remains the authoritative business/scheduling core.
- Normal production actions stay in the same open workbook.
- Production/staging `.xlsm` writes are native Excel/VBA/COM-compatible, never openpyxl saves.
- DAILY_INPUT facts are dated overlays, never destructive edits of the recurring schedule.
- Replacement flow remains: base -> cancellations/no-shows -> absences -> surviving replacements -> global solve -> double-booking review -> explicit resolution -> recompute.
- Final unresolved replacement count must be zero.

## Confirmed native state before V15

The patient all-day DAILY_INPUT branch reached a stable native state before the therapist-absence smoke. MASTER uses the confirmed layout D:J for therapy overlays and M for daily status; K/L remain protected from daily status writes. Non-active dated sessions preserve recurring treatment text and show the dated inactive state through presentation rather than destructive replacement of recurring text.

`THERAPIST_DAILY` is the supervisor workload surface. Therapist workload is intentionally not exposed on MASTER. The supervisor panel has the accepted V14 mint/coral presentation, locked workload badges, restore controls and `Επιστροφή στο MASTER` navigation.

## V14 therapist-absence native findings

A real therapist absence exposed three independent defects:

1. the retired `ΑΝΑΝΕΩΣΗ` control was still visible on `THERAPIST_DAILY`;
2. therapist `Αιτία` had no dropdown/list;
3. after `Εφαρμογή`, Excel reported `Οι αλλαγές αποθηκεύτηκαν, αλλά η κεντρική εικόνα δεν ενημερώθηκε σωστά.` and the flow stopped before orphan-patient replacement handling.

The third finding is now explained exactly. `RefreshOperationalViews()` still required `master.Shapes("ekaWorkload_Refresh")`, but the V14 privacy/UI architecture intentionally removed that MASTER workload control. The daily facts were already saved, then refresh failed, so `ApplyDailyInputPreview()` exited before `ResolveDailyReplacementsInPlace()`.

## V15 remediation now on branch

### Therapist absence reason dropdown

`Python/rehab_excel/daily_input_sheet.py`

- defines `THERAPIST_ABSENCE_REASONS`;
- stores the list on hidden `_PY_LISTS`;
- creates named range `PY_THERAPIST_REASONS`;
- applies it to therapist column E (`Αιτία`).

Current options:

- Άδεια
- Ασθένεια
- Ρεπό
- Εκπαίδευση
- Υπηρεσιακή ανάγκη
- Άλλο

### Operational refresh / replacement continuation

`Python/rehab_excel/daily_input_action_vba.py`

- no longer looks up `MASTER_SCHEDULE.Shapes("ekaWorkload_Refresh")`;
- still calls `RenderTherapistWorkloadStrip` and `ApplyMasterView`;
- after successful refresh, `replacement_sessions_needed > 0` continues into `ResolveDailyReplacementsInPlace()`.

### Retired refresh-control cleanup

`Python/rehab_excel/daily_input_change_navigation_vba.py`

- cleanup remains scoped to `THERAPIST_DAILY`;
- handles Forms buttons, ActiveX/OLE controls and ordinary Shapes;
- exact caption `ΑΝΑΝΕΩΣΗ` is accepted as retired legacy identity even if Excel rewrites the control name;
- qualified macro actions are reduced to their terminal macro name so `RefreshAll` variants are recognized;
- unrelated application controls are not targeted.

### Regression guards

New/updated:

- `Tests/test_therapist_absence_native_flow_regressions.py`
- `Tests/test_therapist_navigation_and_locked_workload.py`

The guards require no dependency on the retired MASTER workload shape, require replacement continuation to remain wired, require therapist-reason dropdown wiring, and cover legacy refresh removal across Excel control types.

## Therapist absence cascade restore checkpoint

The replacement engine remains protected and unchanged. The restore path now handles dependency cleanup around it.

For one restored therapist-absence source row, the preparation step compares the effective therapist absences before and after ignoring only that source row. A session is considered released only when it was covered before and is no longer covered by any therapist absence afterwards. This protects overlapping absence rows.

Only dependencies attached to released sessions are restored:

- ACTIVE rows in `DAILY_REPLACEMENT_FACTS` are changed to `RESTORED`;
- only solver-generated `Αναστολή θεραπείας` rows with the authoritative global-solver reason are cleared;
- unrelated replacement facts, manual suspensions and other therapist absences remain untouched.

The rebuilt daily state excludes those dependent facts before calculating the new MASTER and THERAPIST_DAILY projections. VBA applies dependency mutations, clears the selected source row, applies the validated projection, saves, then refreshes workload/ledger.

The same rollback copy now covers the full DAILY_INPUT mutation range and `DAILY_REPLACEMENT_FACTS`, so any dependency or projection verification failure restores the pre-action state.

Protected replacement behavior remains guarded separately by the established global replacement regression gate. No solver, ranking, policy, adjacency, capacity or scarcity logic was modified for cascade restore.

## Exact resume point

Before native restore testing, run the focused cascade-restore tests together with the protected replacement gate. Only after both are green should a fresh disposable workbook be used to test therapist absence -> replacement -> restore in native Excel.

If restore fails natively, do not click the restore action again. Read `%TEMP%\eka_daily_input_restore_last_error.log` immediately and diagnose the recorded phase.

## Known follow-up debt after native restore is green

- Guided double-booking VBA contains historical raw-string/backslash risk. Normalize only the path literals if the native flow reaches and exposes that defect; do not blanket-replace JSON escape handling.
- MASTER daily presentation reset should eventually preserve intentional recurring font colors rather than using a blanket black reset.
- Python 3.14/openpyxl `ZipFile.__del__` warnings remain non-business cleanup debt in disposable tests.

## Safety reminders

- baseline immutable;
- rule -> test -> Python -> green -> Excel integration;
- production action contract: validate -> authoritative compute -> in-place native write -> semantic read-back -> Save -> continue in application;
- no production close/reopen, live-file replacement, visible console or file-lock waiting workflow;
- diagnostic-first native testing, no blind retry;
- Windows tests use `python -m pytest`;
- documentation updates are part of implementation.
