# H7 BC install passed; live UI pending | 2026-11-03

## Operator-confirmed BC install gate

Source: `Excel/previews/UNIFIED_PATIENT_HISTORY_H6_20261103_AZ.xlsm` (GREEN/FROZEN).
Fresh candidate: `Excel/previews/UNIFIED_PATIENT_HISTORY_H7_20261103_BC.xlsm`.

Operator output:
```text
BC COPY VERIFIED
H7 PATIENT HISTORY UI INSTALL OK
Entry: DAILY_INPUT -> Στατιστικά ασθενών
Technical history sheets: veryHidden
Persisted VBA reopen verification: OK
```

**Status: BC INSTALLATION GREEN; native live UI NOT YET PROVEN.** Do not declare BC GREEN/FROZEN until the operator clicks and verifies all report modes. Source tests earlier: 12 passed in 3.27s.

## Immediate native click gate

Open BC (no edits to frozen source AZ), navigate DAILY_INPUT -> `Στατιστικά ασθενών`.
Switch:
1. Daily, service date 2026-10-30.
2. Weekly, 2026-10-26 through 2026-11-01.
3. Monthly, October 2026.
4. Overall.
5. Daily again: ensure button row persists.
6. Back to MASTER.

Expected monthly report: 77 grouped patient rows with Scheduled=80, Performed=77, PatientAbsences=3, Cancellations=0, Replacements=0, based on prior read-only TSV proof on AZ.

BC inherits old AZ toolbar layout and may have old DAILY_INPUT close/save controls or overlap on MASTER. These are not new H7 defects. BB is final H6 layout GREEN/FROZEN and is not to be modified. Later integrate H6 final UI + H7 on a **fresh** candidate.

If live UI works, mark BC GREEN/FROZEN for the H7 isolated UI proof and record result. If any runtime error, missing controls, wrong counts, or bad navigation, freeze BC FAILED/FROZEN, capture exact failure; fix source and use new suffix. Never repair BC. Preserve production and all prior green artifacts.
