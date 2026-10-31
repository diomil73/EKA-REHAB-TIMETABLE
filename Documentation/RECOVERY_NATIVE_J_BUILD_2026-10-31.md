# Recovery native J build — 2026-10-31

## Build result

Fresh artifact:
`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_J.xlsm`

Build completed successfully.

Evidence from local build:
- unified preview stages all reported `[OK]`
- source workbook remained unchanged
- VBA project present
- DAILY_INPUT present
- application contract verified
- application contract required component count: 24
- artifact size: 471106 bytes
- artifact last write: 2026-10-31 17:54:42 local time
- SHA256 begins `3845F670CBB8CC8FD0C5396C8192FEE6D208204CC76...`

## Native gate for J

Open only J.

Primary row-restore cycle:
1. Mark one patient absent in DAILY_INPUT.
2. Apply DAILY_INPUT.
3. MASTER must keep the treatment visible with red strike-through.
4. MASTER reason/status must be red and readable, not struck through.
5. THERAPIST_DAILY must render the change ledger below the operational grid.
6. The active ledger row must offer `Επαναφορά`.
7. Restore that row.
8. Patient must return to normal in THERAPIST_DAILY and MASTER.
9. The corresponding DAILY_INPUT source row must be cleared.
10. The ledger row must remain visible as history with status `Επαναφέρθηκε`.
11. The same J workbook must remain open; no close/reopen path is permitted.

If row restore passes, continue in the same J with selected restore and restore-all.

Therapist-absence restore remains intentionally disabled until coordinated rollback of related replacement facts is implemented and tested.

Do not run CloseOperationalDay until this restore gate passes.
