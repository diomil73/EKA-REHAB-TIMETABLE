# V14 Supervisor UI Checkpoint

Status: native visual checkpoint confirmed by user on 2026-10-27.

## Confirmed V14 state

`DAILY_INPUT_IN_PLACE_V14.xlsm` is the current visual/native reference for `THERAPIST_DAILY`.

Confirmed:

- the lower supervisory panel is visibly redesigned and accepted as clearly better;
- the panel uses mint background/header treatment and soft coral restore actions;
- `Επιστροφή στο MASTER`, `Επαναφορά επιλεγμένων` and `Επαναφορά όλων` are grouped in the lower management area;
- the old top-right `ΑΝΑΝΕΩΣΗ` control was manually removed from V14 and the workbook was saved;
- the old `RefreshAll` VBA procedure itself was intentionally retained because deleting the obsolete button does not imply deleting the macro.

## Legacy refresh-button retirement

Future preview builds must not recreate the obsolete top-right refresh control.

The installer now removes it only when all three identifiers match:

- shape/form-control name: `Button 1`;
- visible caption: `ΑΝΑΝΕΩΣΗ`;
- assigned macro: `RefreshAll`.

This exact-match rule is intentional so unrelated controls are not deleted.

Relevant implementation:

- `Python/rehab_excel/daily_input_change_navigation_vba.py`
- `Tests/test_therapist_navigation_and_locked_workload.py`

## Privacy contract retained

Therapist workload remains absent from the general `MASTER_SCHEDULE` surface. `THERAPIST_DAILY` is a supervisory Physiotherapy surface intended for `SUPER_USER` and a future explicitly authorized Physiotherapy Head.

## Next native functional checkpoint

The next native branch to verify is therapist absence / global replacement continuation.

Target sequence:

`DAILY_INPUT therapist absence -> authoritative recomputation -> replacement continuation -> optional guided double-booking review -> in-place persistence/read-back -> refreshed MASTER/THERAPIST_DAILY/ledger`

No production action may close/reopen the live workbook.

## Known targeted risk before/within guided double-booking smoke

`Python/rehab_excel/double_booking_review_vba.py` still contains some Python-raw VBA path literals with doubled backslashes (`\\`). VBA does not use backslash escaping for filesystem separators. These path literals must be normalized deliberately if the guided double-booking runtime path is exercised; do not blanket-replace all backslashes because the same module also contains intentional JSON escaping logic.
