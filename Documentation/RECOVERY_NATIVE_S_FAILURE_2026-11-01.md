# Native S failure evidence — 2026-11-01

Artifact: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_S.xlsm`

Build path: canonical `Python/tools/build_unified_preview.py` only. No recovery wrapper.

Build evidence:
- unified build completed successfully,
- source workbook unchanged,
- application contract verified with 24 required components,
- artifact size reported locally as 472583 bytes,
- SHA256 output began `E7AEAB0BEDE823CB78312C9437331686212EF730A03...`.

Native patient-absence smoke result:
- MASTER treatment presentation was correct after Apply: recurring treatment remained visible, red and struck through. This confirms the character-level D8 strike/read-back fix reached the built workbook.
- MASTER patient status/reason presentation remained wrong, matching the previous behavior. The reason/status text was still struck through instead of red/no-strike.
- THERAPIST_DAILY behavior remained wrong, matching the previous destructive/overwrite behavior rather than preserving existing operational details.
- Restore produced the same two-stage failure messages as before and made no effective restoration.

Interpretation:
- S is not a total build miss because the D8 fix is visibly present.
- The remaining status, THERAPIST_DAILY and restore failures must not be treated as proof that the wrong workbook was opened.
- Do not build another artifact until the final installed VBA in S and the current error logs are inspected read-only.

Next diagnostic:
- `Python/tools/inspect_native_daily_restore_contract.py` opens S read-only and reports installed semantic markers in `modDailyInputAction` and `modDailyChangeLedger`, plus `%TEMP%` DAILY_INPUT/restore error logs.

S is frozen as evidence and must not be modified or reused for repair-over-repair.
