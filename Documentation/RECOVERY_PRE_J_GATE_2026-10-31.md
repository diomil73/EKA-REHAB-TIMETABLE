# Recovery pre-J gate — 2026-10-31

Focused post-I verification gate passed: 37 passed in 4.00s.

I build name is retired. The I build failed closed during operator visibility post-install verification because Excel COM returned CRLF line endings while the verification compared LF literals. No I artifact was created and the immutable source workbook hash remained unchanged.

The bounded fix normalizes installed VBA source before semantic verification. Runtime MASTER patient-status behavior is unchanged: the absence reason is red and explicitly not struck through.

Next artifact: fresh `UNIFIED_CLOSE_DAY_RECOVERY_20261031_J.xlsm` only. Expected application contract: 24 components plus required `OUTPATIENT_SCHEDULE` and `DAILY_CHANGE_LEDGER` sheets.

Native J smoke order:
1. patient absence applies;
2. MASTER treatment remains visible with red strike-through;
3. MASTER reason/status remains red and not struck through;
4. change appears in ledger below THERAPIST_DAILY;
5. row restore clears the authoritative DAILY_INPUT source row, recomputes THERAPIST_DAILY and MASTER in-session, and leaves ledger history marked `Επαναφέρθηκε`;
6. only after row restore passes, test selected/all restore;
7. therapist-absence restore remains intentionally blocked until coordinated replacement-fact rollback is implemented and proven.
