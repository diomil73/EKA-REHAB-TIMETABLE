# Project handoff: exact next steps, checks and stop conditions

**Branch:** `work/history-statistics-rebuild`  
**Resume position:** H7 statistics, after **15 passed** source gate and **successful read-only monthly AZ report**. Do **not** ask the user to rerun either unless a new code change makes it necessary.

## Immediate next development task

Build out H7 read-only reporting for users, using the already verified `CURRENT_FINAL` patient, therapist, and snapshot readers. Decide the next narrow source slice after inspecting code, then add source tests. Do not duplicate H4's already GREEN therapist-history UI. Intended next slices:

1. Review the current H4 user-facing `THERAPIST_HISTORY_REPORT` and report-generation tools; reuse them rather than creating redundant therapist reports.
2. Add patient-facing history/statistics presentation and filters for day, ISO week, month and overall period, including reason/cancellation-kind semantics. User-facing names and controls must be separate from hidden technical storage.
3. Audit daily/therapist/patient reconciliation from current snapshot only, with explicit fail-closed behavior for gaps, duplicates, mismatched dates and versions. Avoid assuming the sum of therapist active sessions equals patient delivered sessions when student replacements or shared-session rules alter ownership.
4. Add native UI only via new disposable smoke workbook and complete source tests before asking for native proof. No production workbook modification.
5. Preserve H4 report behavior and all H2–H6 native evidence; test that new navigation, session save and protection are not regressed.

## Canonical command: read-only monthly report, ALREADY GREEN

```powershell
python .\Python\tools\report_period_statistics.py `
  --workbook ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm" `
  --period month `
  --start 2026-10-01 `
  --end 2026-10-31
```

**Observed user output:** `ok=true`, October 2026 `finalized_days=1`, `scheduled=80`, `performed=77`, `patient_absences=3`, `cancellations=0`, `replacements=0`, `workbook_unchanged=true`. Already tested; reference evidence only.

## Source regression gate (ALREADY GREEN)

```powershell
git pull --ff-only origin work/history-statistics-rebuild

python -m pytest `
  Tests/test_patient_statistics_h7.py `
  Tests/test_historical_statistics_h7.py `
  Tests/test_period_statistics_h7.py -q
```
Result supplied by user: `15 passed in 3.29s`. Re-run only after relevant changes or for a later explicit regression gate.

## Inspectors and native tests when actually necessary

```powershell
python .\Python\tools\inspect_close_day_native_outcome.py `
  --workbook ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H5_20261103_AX.xlsm"

python .\Python\tools\inspect_day_amendment_outcome.py `
  --workbook ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H5_20261103_AY.xlsm"

python .\Python\tools\inspect_unified_day_save_controls.py `
  ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H6_20261103_BB.xlsm"
```
These are read-only reference inspections, **not a request to repeat known green proofs**.

## Branching rules for every future gate

- **Source tests GREEN:** record exact count/time, commit and checkpoint; proceed to next independent slice proactively.
- **Source tests FAIL:** read smallest relevant traceback, make surgical source fix with a regression, run targeted gate; never assume a failure is an Excel native problem.
- **Read-only report GREEN:** check row totals, period boundaries, `CURRENT_FINAL` semantics and no unexpected duplicate count; record factual output.
- **Read-only report FAIL / inconsistent:** do not write, repair or convert the frozen workbook. Diagnose reader and compare with source history semantics, add regression, then retest.
- **Native workbook build/install fails:** stop before opening it, keep diagnostic information; use fresh suffix for next attempt. Never reuse a failed workbook.
- **Native test succeeds:** inspect authoritative output; mark artifact GREEN/FROZEN, do not change it.
- **Native test fails:** record full error, freeze artifact FAILED/FROZEN and inspect diagnostics read-only; no repair in place.
- **Unknown/missing evidence:** mark unverified; do not invent tests or infer a success from UI appearance alone.

## Release blockers and deferred UX

**Mandatory before production release:** user observed arbitrary editable cells across workbook. Implement fine-grained cell/sheet protection and macro-appropriate permission refresh. Require native proof that attempted invalid user edits are blocked while intended absence entry, apply, single Save Day first close and same-session v1→v2 remain usable; preserve technical `veryHidden` sheets. Details: `Documentation/BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md`.

Cosmetic underline in `διακομισμένος` is intentionally deferred. No other request to spend time on it.

## Quick commands for reconnect

```powershell
git status --short
git branch --show-current
git pull --ff-only origin work/history-statistics-rebuild
```

If `git pull` refuses due to local changes: **do not reset or force-pull**; inspect `git status` and preserve local work. If Excel is open during a COM installer, close Excel first. Report readers using openpyxl are read-only.

**Do not tell the user to wait for background results.** Complete each achievable step in the current turn, commit reliable source/doc artifacts, and keep checkpoint docs updated.
