# H7 historical day semantics and previous-screen navigation | 2026-11-04

## User clarification

User asked whether "Ημέρες" values 1 and 0 in patient statistics are based only on accumulated historical data and requested navigation directly to the immediately previous page from both the patient and therapist report screens, without going via MASTER.

Answer:
- `Ημέρες = 1` means that, **within the chosen report period**, that patient has at least one scheduled/finalized session on one distinct service date of the CURRENT_FINAL historical data. It does not necessarily imply a performed treatment: a missed or cancelled session still counts toward the presence of a finalized session-day.
- `Ημέρες = 0` is the explicit zero-activity row for a patient who is in the **current** PATIENTS registry but has no CURRENT_FINAL session in that period.
- More finalized dates can produce 2,3,...; superseded snapshots do not double-count; missing days never inferred.
- Current registry is an overlay, **not an immutable historic registry** at prior dates.

## Source navigation solution staged

The entry to both report screens is THERAPIST_DAILY; therefore their previous page is the same shared entry screen. The existing proven MASTER toolbar VBA macro `ToolbarTherapistDaily` already activates this sheet and calls `KeepApplicationShell`, no new duplicate navigation macro needed.

- `Python/rehab_excel/patient_history_report_vba.py` commit `099de66`: report Back button now says `← Προηγούμενη` and executes `ToolbarTherapistDaily`.
- `Python/rehab_excel/therapist_history_report_vba.py` commit `7bfe93d`: same correction for therapist report, in both runtime button creation and native installer, with 116-point width to fit caption; does not otherwise change H4 report aggregation.
- `Tests/test_h7_history_previous_screen_navigation.py` commit `b091ba8`: three focused tests to enforce common back destination, runtime and installation consistency.

These source edits are **not yet tested by operator** and **not installed in native BE**. Do not patch BE or previous frozen files. Next native candidate requires a fresh unique suffix and must also consider unresolved `ΑΝΑΝΕΩΣΗ` shape (see `Documentation/CHECKPOINT_H7_BE_REFRESH_BUTTON_AND_TOOLBAR_2026-11-04.md`).

## Next source gate

```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_h7_history_previous_screen_navigation.py Tests/test_therapist_history_ui.py Tests/test_patient_history_ui_h7.py -q
```
Expected 18 passed (3+10+5), pending operator output.

## Outstanding

- BE screenshot: stray `ΑΝΑΝΕΩΣΗ` far right and stats buttons below restore area, not aligned with navigation. Source toolbar row fix is staged; no native verification. Read-only `Python/tools/audit_h7_therapist_daily_shapes.py` on frozen BE pending to determine refresh macro; do not blindly remove.
- Historic patient totals: one finalized day, 80 scheduled, 77 performed, 3 absences. Full-registry overlay 93 patients with 16 zero-activity rows is source-tested 19 passed in 3.46s, but **BE native report outcome** not yet explicitly confirmed for all modes.
- BE is not accepted GREEN/FROZEN. Preserve BE diagnostic, AZ/BB/AQ GREEN/FROZEN; production immutable. Final worksheet protection remains mandatory release gate.
