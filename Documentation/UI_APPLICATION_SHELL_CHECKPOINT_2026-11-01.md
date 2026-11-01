# UI / Application Shell Checkpoint — 2026-11-01

## Confirmed product direction

The workbook is an implementation engine, not the intended visual identity of the product.

The user-facing experience should evolve toward a professional application shell and should not look or behave like a collection of Excel worksheets with ad-hoc buttons.

MASTER_SCHEDULE remains a clean, read-only operational hub in normal-user mode.

Normal-user MASTER must not contain operational/admin toolbar buttons. A future super-user mode may expose privileged controls, but only as an explicit separate mode. Do not use hidden or accidental MASTER controls as a substitute for proper privileged UX.

Operational actions belong on the relevant workflow surface or in a dedicated application navigation/menu layer. The current footer buttons are transitional scaffolding, not the final UI design.

The user recalled prior research/benchmark discussion of comparable software and an agreed direction toward a polished application-like presentation. Archived project material exists and should be consulted before the later full UI redesign instead of starting from scratch.

## H3 terminology

User-facing label for the current closed-day persistence/version-chain operation: `Αποθήκευση ημέρας`.

Internal implementation terms such as amendment/version-chain may remain in Python/VBA names and technical documentation, but they should not leak into normal-user wording.

## Native UI artifact state

`Excel/previews/UNIFIED_DAY_SAVE_UI_20261101_AD.xlsm` is frozen failed UI-layout evidence.

Observed from operator screenshot:

- legacy MASTER buttons were still visible: `Διαχείριση Μητρώων`, `Απουσίες / Ακυρώσεις`, `Ημερήσιο πρόγραμμα`, `Save & Exit`,
- expected clean MASTER contract was therefore not applied,
- expected `Αποθήκευση ημέρας` placement was not visible in the inspected UI,
- no business-operation click should be used as proof from AD.

Root cause: the dedicated day-save UI installer still contained stale verification/installation assumptions from the earlier MASTER-button prototype. It installed only the day-amendment action and verified a MASTER button instead of reapplying the current clean-MASTER shell and verifying DAILY_INPUT placement.

Fixes:

- `71ce7db` — day-save installer reapplies `install_master_toolbar(...)`, verifies no `ekaToolbar_*` shapes remain on MASTER, and verifies the day-save button on DAILY_INPUT.
- `4736005` — regression tests for clean MASTER and DAILY_INPUT placement.

## Next proof

Run focused source tests, then use a fresh uniquely named AE workbook. Do not reuse or repair AD.

The immediate H3 proof is about safe business workflow integration, not final visual design. After H3 functionality is proven, return to the previously researched application-shell/UI direction and perform the broader professional UI redesign as a distinct product slice.
