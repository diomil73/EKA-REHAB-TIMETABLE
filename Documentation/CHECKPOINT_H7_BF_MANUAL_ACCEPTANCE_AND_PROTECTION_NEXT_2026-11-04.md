# H7 BF manual acceptance green; worksheet protection audit next | 2026-11-04

## H7 final functional acceptance

Operator confirmed: **"ολα γινανε σωστα"** after the BF native manual test sequence.

This confirms user acceptance of:
- integrated THERAPIST_DAILY toolbar placement,
- retained and restyled `ΑΝΑΝΕΩΣΗ` control,
- RefreshAll execution,
- patient statistics entry/report modes,
- therapist statistics entry/report,
- `← Προηγούμενη` direct return to THERAPIST_DAILY.

Supporting evidence:
- Focused source gate: `26 passed in 4.87s`.
- Fresh BF: `UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm`.
- Independent persisted read-only audit: `H7 INTEGRATED AUDIT OK; workbook not changed`.
- Toolbar coordinates: RefreshAll x=170,y=563.6,w=92; therapist stats x=270,y=563.6; patient stats x=424,y=563.6.
- VBA Module3, modUnifiedDaySave, modTherapistHistoryReport, modPatientHistoryReport present.
- Both reports visible; technical history sheets veryHidden.
- User manually confirmed all functions worked correctly.

**H7 is functionally accepted.** BF remains a disposable smoke workbook, not production. Do not patch/reuse BF or alter production.

## Next milestone: final worksheet protection hardening

The user previously observed that free text could be entered in many worksheet cells. This is now the next active milestone after H7 completion.

Read backlog:
`Documentation/BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md`.

Acceptance:
- MASTER_SCHEDULE read-only.
- DAILY_INPUT only intended editable input fields and sanctioned workflows.
- THERAPIST_DAILY and operational views shield calculated/projection cells.
- Registries and underlying records editable only through sanctioned forms/macros.
- History sheets remain veryHidden and protected; no direct edits.
- Protect workbook structure and control worksheet selection appropriately.
- Re-protect on open/session changes/save/rollback/error; do not assume UserInterfaceOnly persists.
- No broad unprotect windows.
- Prove arbitrary grid edits blocked while intended entry, first close, amendment v1→v2, refresh and statistics remain functional.
- No blanket protection until allowed edit zones are mapped.

## Read-only baseline diagnostic added

- `Python/tools/audit_workbook_protection.py` commit `aa0cd3c`, permission reporting enhancement `cf04e93`.
- `Tests/test_workbook_protection_audit.py` commit `88d9446`.
- Diagnostic inventories workbook structure/windows protection; per-sheet ProtectContents, ProtectDrawingObjects, ProtectScenarios; EnableSelection; UsedRange and Locked status; protection allowances; unlocked ranges on mixed locked/protected sheets.
- Opens workbook with `ReadOnly=True`, `UpdateLinks=0`; closes with `SaveChanges=False`; never calls Protect, Unprotect or Save.
- If a protected sheet's mixed UsedRange exceeds 150,000 cells, scan is marked incomplete and returns nonzero rather than implying full coverage.

## Next operator action

After pull and source tests, close Excel and run the audit against accepted BF:
```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

python -m pytest Tests/test_workbook_protection_audit.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Protection audit tests failed" }

python .\Python\tools\audit_workbook_protection.py `
  ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261104_BF.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Protection inventory incomplete" }
```

Expected: 5 tests passed. The native audit is read-only. Share its full output.

## Protection implementation is NOT started

Do not yet call any sheet Protect/Unprotect, lock/unlock ranges, or add protection passwords. First review the BF audit output and map intended writable ranges per sheet and all macro/form write zones. Then design focused source tests and fresh next-candidate native smoke. Preserve BF as green evidence.

## Protection audit diagnostic formatting fix — 2026-11-04

First operator run exposed one source-test assertion mismatch: the audit emitted the locked state through an f-string label, while the test required an explicit `"UsedRange.Locked"` literal. The audit itself already read the property. Corrected by printing `UsedRange.Locked:` as a dedicated label. Also moved protection allowance reporting before the early unprotected-sheet exit, so formatting/filter/sort/insertion allowances are inventoried for every worksheet, not only mixed-lock protected sheets.

Fix commit: `1d65d90`.

No Excel workbook was changed by this fix. The operator had issued the read-only audit command after the failing test but did not include its output in the report, so treat that audit result as unreviewed. Pull the fix, rerun the five focused tests, then rerun the read-only audit against BF and share the complete output. No protection changes are authorized yet.

## Operator protection inventory received — 2026-11-04

The user supplied two complete read-only audits of BF. Both report `WORKBOOK PROTECTION AUDIT COMPLETE`, `ReadOnly: True`, and `Workbook unchanged: True`.

Key findings (23 worksheets):
- Workbook structure and windows are not protected: ProtectStructure=False, ProtectWindows=False.
- Unprotected worksheets: `_PY_LISTS`, `DAILY_INPUT`, `PATIENT_PLANNER`, `DAILY_CHANGE_LEDGER`, `OUTPATIENT_SCHEDULE`, `MASTER_SCHEDULE`, `THERAPIST_ATTENDANCE`, `REPLACEMENTS`, `REPLACEMENT_LOG`, `DAILY_LOG`, `STATISTICS`, `PATIENTS`, `NEW_PATIENT`, `SETTINGS`, `SESSIONS`, `CONFLICT_LOG`, `THERAPIST_HISTORY_REPORT`, `THERAPIST_DAILY`.
- Already protected: `PATIENT_HISTORY_REPORT`, `FINALIZED_DAY_SNAPSHOTS`, `THERAPIST_DAY_HISTORY`, `PATIENT_SESSION_HISTORY`, `DAY_HISTORY_EVENTS`.
- The four technical history sheets are veryHidden and protected. Patient history report is visible and protected. Therapist history report is visible but unprotected.
- All reported UsedRange.Locked values are True. This does NOT prevent editing when ProtectContents=False; worksheet protection is the missing enforcement layer.
- EnableSelection is xlNoRestrictions on every worksheet, including protected sheets. This is not by itself an edit vulnerability if protection is active, but should be reviewed for technical sheets.
- All displayed protection allowances are False on protected sheets.
- The read-only scan completed; no scan errors/truncation.

The test still failed after first formatting fix because source contract searches exact literal `"UsedRange.Locked"`; implementation printed `"UsedRange.Locked:"` with colon inside the string, so assertion remained false. Corrected to `print("UsedRange.Locked", locked_label)`. Fix commit: `e89e36f`. Need pull and rerun the five tests; no native workbook rerun is needed solely for this output-label-only source fix, but do rerun read-only inventory after tests if desired.

No protection implementation has been made. Before locking, inspect all VBA/Python write paths and map sanctioned editable inputs and write targets. In particular do not infer that DAILY_INPUT can be wholly locked: its forms/macros may write cells, and manual edit zones must be explicitly mapped. Also inspect THERAPIST_DAILY controls/data writes and report regeneration. Preserve the five currently protected sheets as baseline.

## Operator test confirmation and next VBA audit

Operator reports `5 passed in 0.64s` for `Tests/test_workbook_protection_audit.py`. The diagnostic formatting contract is now green.

Tracked source map: `Documentation/WORKSHEET_PROTECTION_WRITE_PATH_MAP_2026-11-04.md`. It records known sanctioned writes and highlights unresolved workbook-only macro paths, external COM writer implications, and dynamic sheet creation. No protection implementation has been started.

Next add-on diagnostic is the read-only VBA inventory. Run the combined eight tests and `audit_vba_write_surfaces.py` as specified in `Documentation/CURRENT_HANDOFF.md`. Only after this inventory is reconciled can per-sheet policy be finalized.
