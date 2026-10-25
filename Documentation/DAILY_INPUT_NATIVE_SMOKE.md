# DAILY_INPUT in-place native smoke

Last updated: 2026-10-25

## Native checkpoint

Focused regression before smoke: `57 passed in 2.65s`.

Native workbook: disposable application-shell preview built from `Excel/previews/DOUBLE_BOOKING_REJECTED.xlsm`.

Observed result on the first in-place DAILY_INPUT smoke:

- no console/Command Prompt window appeared;
- the live workbook stayed open;
- the action failed closed during in-place apply/read-back;
- operator message reported that the daily state was not applied and changes were rolled back;
- rollback visibly returned the workbook to its previous state.

This confirms the rollback safety path works natively. The failure is not related to the retired close/reopen transaction model.

## Finding and fix

Rich-text DAILY_INPUT patches can contain mixed character formatting. In native Excel, a whole-cell property such as `Range.Font.Size` may be `Null` when character formatting is mixed. The first VBA bridge attempted whole-cell minimum-font conversion/verification, which can therefore trigger a false runtime failure and rollback.

`Python/tools/prepare_daily_input_in_place.py` now omits whole-cell `min_font_size` from serialized payloads when a patch contains rich-text runs. Rich-text formatting continues to be verified at the individual character-run level. Plain patches keep minimum-font serialization.

Regression coverage in `Tests/test_prepare_daily_input_in_place.py` now checks this mixed-font contract.

## Next gate

Build a fresh disposable application-shell workbook from `DOUBLE_BOOKING_REJECTED.xlsm` and repeat the same DAILY_INPUT native smoke. Success means the workbook remains open, no console appears, presentation is applied, and routing goes to replacements only when unresolved therapist-absence work remains.
