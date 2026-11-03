# H7 BC visual feedback and final H6+H4+H7 integrated UX staging | 2026-11-03

## BC status

- BC original source copy check and H7 installer/persisted readback GREEN.
- Operator viewed H7 patient report and said "καλά φαίνονται αλλά θα φανεί στη χρήση". Do **not** promote BC to full integrated acceptance.
- Three native UI observations:
  1. MASTER `Αποθήκευση ημέρας` returned to wrong old position. Expected: H6 BB layout, immediately after Save & Exit and no overlap.
  2. `Στατιστικά ασθενών` on DAILY_INPUT is misplaced and should not be there. Move next to `Στατιστικά θεραπευτών`. Therapist button not visible in BC because H4 installer was never applied to the AZ-derived test data.
  3. Patient IDs appeared lexicographically (1, 10, ..., 2, 20). Show 01, 02, ... 10, 20 in natural numeric ordering.

BC remains **visual provisional / superseded by feedback**, FROZEN; never repair it. Preserve AZ and BB GREEN/FROZEN.

## Source changes

- `Python/rehab_excel/patient_history_reports.py` commit `4131fc8`: natural numeric patient sort by integer, without altering actual historical identifiers.
- `Python/tools/export_patient_history_ui.py` commit `2ceed32`: pad *single-digit all-numeric* ID **only in display output**, never in historical storage; preserve prefixed IDs and existing zero-padded values.
- `Python/rehab_excel/patient_history_report_vba.py` commits `62488e5`, `8fc495d`: patient statistics navigation now on THERAPIST_DAILY aligned just right of H4 therapist statistics, delete legacy DAILY_INPUT nav, mark patient ID display column as Excel text to preserve 01.
- `Python/tools/install_patient_history_ui.py` commit `8b11df3`: on persisted reopen, enforce both buttons exist on THERAPIST_DAILY, are aligned and do not overlap, and none remains on DAILY_INPUT.
- `Tests/test_h7_integrated_toolbar_and_patient_order.py` commit `c12f226`: 7 targeted new tests (two additional smoke-builder guards in commit b14d56a).

## Next SOURCE gate, not yet run

```powershell
git pull --ff-only origin work/history-statistics-rebuild
python -m pytest Tests/test_patient_history_ui_h7.py Tests/test_patient_history_reports_h7.py Tests/test_h7_integrated_toolbar_and_patient_order.py -q
```

Expected 19 passed if new tests pass. **Do not claim green until user supplies output.**

## Planned integrated native gate after source green

Fresh unique **BD**, copied byte-for-byte from frozen AZ (which carries the H7 real CURRENT_FINAL 80/77/3 history). Close Excel first. Do not change AZ, BB, BC, production.

Run these existing installers **in order on BD only**, with a fail-fast stop after each:
1. `Python/tools/install_unified_day_save_ui.py` to restore MASTER-only BB control layout (removes old DAILY_INPUT close/amend buttons).
2. `Python/tools/install_therapist_history_ui.py` to add H4 therapist history report and its `THERAPIST_DAILY` statistics button.
3. `Python/tools/install_patient_history_ui.py` to add H7 patient report beside it with reopen/placement verification.

Then inspect controls read-only with `Python/tools/inspect_unified_day_save_controls.py`, and perform **focused integrated visual checks only**:
- MASTER one Save Day after Save & Exit, no overlap.
- DAILY_INPUT no save/close buttons, no patient statistics button.
- THERAPIST_DAILY two aligned stats buttons.
- Patient daily, weekly, monthly, overall, back, natural 01/02/10/20 numbering; unchanged October totals 80 scheduled, 77 performed, 3 absences. H4 therapist report opens and keeps its buttons.
- History technical sheets veryHidden; no authoritative save/amendment rerun.
- Native failure: freeze BD FAILED/FROZEN, record exact evidence, source fix and new suffix. Native success: mark BD GREEN/FROZEN, checkpoint.

## Release blocker

Final worksheet editing protection remains deferred but mandatory before production, per `Documentation/BACKLOG_FINAL_WORKSHEET_PROTECTION_2026-11-03.md`.

## Single fail-fast integrated installer added

`Python/tools/prepare_h7_integrated_ui_smoke.py` commit `96eea88`, regression `b14d56a`. Automatically copies frozen AZ to a fresh non-existing BD, verifies SHA-256 byte equality, then applies H6 MASTER-only, H4 therapist UI, and H7 patient UI installers sequentially with nonzero failure stops. Uses separate Excel COM installer runs. No claim of Windows/native execution until operator output.

After source tests GREEN and Excel fully closed:
```powershell
python .\\Python\\tools\\prepare_h7_integrated_ui_smoke.py --output ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: BD integrated preparation failed" }
python .\\Python\\tools\\inspect_unified_day_save_controls.py ".\\Excel\\previews\\UNIFIED_PATIENT_HISTORY_H7_20261103_BD.xlsm"
```

Verify installer output and read-only audit before opening. Existing BD must fail (never overwrite/reuse). If successful, perform only focused UI checks and freeze.
