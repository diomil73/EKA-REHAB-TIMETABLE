# H7 current registry overlay staged after 93-vs-77 read-only proof | 2026-11-04

## Operator evidence
BD read-only audit:
- 93 PATIENTS registry entries; 80 CURRENT_FINAL sessions; 77 distinct active patient IDs.
- 77 daily and 77 weekly historical activity-report rows.
- 15,19,20,21: in current registry; no CURRENT_FINAL sessions for that date.
- 18,22: in registry and both report-data sets; apparent Excel UI omission needs visibility checking.
- Registry IDs with no finalized sessions: 02,06,15,19,20,21,24,25,30,39,41,50,64,69,73,75.
- No finalized numeric patient ID absent from registry. Source audit marker: H7 REGISTRY AUDIT COMPLETE.
- The preceding six focused source tests passed in 1.78s. **The new tests below are not yet operator-verified.**

## New source staged
- `Python/tools/inspect_h7_patient_report_visibility.py` `c9d830b`: nonmutating Windows Excel COM view inspection of saved report worksheet, row positions/hidden heights/filters for 15 and 18–22. Does not save or install.
- `Python/rehab_excel/patient_history_reports.py` `e25623c`, `f7812e8`: explicit opt-in overlay of the **current PATIENTS registry** per existing finalized report period, with zero counts when no CURRENT_FINAL sessions; refuses duplicate normalized IDs, nonregistered finalized IDs; correct unique overall grouping across differing patient date spans; does not create new history or dates.
- `Python/tools/export_patient_history_ui.py` `2c2592d`, `37a6747`: opt-in `--include-registered-zero`; default CLI historical activity-only remains unchanged.
- `Python/rehab_excel/patient_history_report_vba.py` `ab6c53a`: for the next freshly generated UI, show current registry zero rows by passing this option; label explains `Τρέχον μητρώο: χωρίς συνεδρία = 0`. Historical SOURCE numbers unaffected. This UI change has **not** been installed in BD.
- `Tests/test_h7_registry_zero_overlay_and_visibility.py` `43ccd28`: eight new tests for zero overlay, multi-day overall, preservation of totals, zero-date refusal, fail-closed duplicates/registry mismatch, exporter default vs opt-in, numeric view recognition, VBA label. **Not yet run by operator**.

## Required next source gate

```powershell
git pull --ff-only origin work/history-statistics-rebuild
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }
python -m pytest Tests/test_h7_registry_zero_overlay_and_visibility.py Tests/test_patient_history_reports_h7.py Tests/test_h7_patient_mode_switch_and_toolbar_fix.py -q
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: H7 new tests failed" }
```

Expected **18 passed** (8 + 7 + 3), subject to proof.

## Diagnostic independent native view audit, no writes to BD

With Excel closed:
```powershell
python .\Python\tools\inspect_h7_patient_report_visibility.py ".\Excel\previews\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
```

This reads **saved** report sheet state, not transient unsaved scroll positions. If IDs 18/22 are present and not hidden, investigate scroll/freeze/viewport vs displayed ordering. If hidden, separate UI view fixes are needed. Never edit BD.

## Future fresh BE acceptance

Only after source gate and, if necessary, diagnosis of 18/22, create new uniquely named BE from frozen AZ via the existing ordered H6+H4+H7 one-command installer. BC/BD must not be overwritten or repaired. Native test: full registry 93 rows incl 16 zero-activity, preserved 80/77/3 session totals, natural patient numbering, visible and correctly aligned/black-label companion buttons, repeated daily/week/month/overall switches (prior intermittent failure), technical history veryHidden. Actual newly installed UI must be tested; source test does not prove it. Preserve protection hardening as a release gate.

## Notes

Current registry membership is not an authoritative reconstructed roster for dates in the past. Mark zero rows as present-day registry overlay, not missed treatments. Do not invent absent days or sessions. Production immutable; prior green/frozen snapshots unchanged.
