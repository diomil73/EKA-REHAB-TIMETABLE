# Recovery native L build — 2026-10-31

## Artifact

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_L.xlsm`

## Operator evidence

The canonical unified preview build completed successfully after the K ledger-refresh installer fix.

Observed build results:
- patient edit flow OK,
- outpatient schedule form OK,
- outpatient conflict guard OK,
- DAILY_INPUT sheet OK,
- MASTER projection refresh OK for 93 patients,
- final user navigation OK,
- MASTER visual baseline OK,
- unified preview reported `UNIFIED PREVIEW OK`,
- source workbook remained unchanged,
- VBA project present,
- DAILY_INPUT sheet present,
- application contract verified,
- independent contract check reported 24 required components,
- artifact SHA-256 begins `E10193F84BC080E968C03935A89FF8694321F781C82...`,
- artifact size 470409 bytes,
- artifact last-write 2026-10-31 18:28:52 local.

## Safety status

- production workbook remains immutable,
- K remains a retired failed-build name,
- J remains frozen forensic evidence,
- L is the current native candidate,
- no close-day smoke yet,
- no therapist-absence restore yet.

## Next native gate

Open only L and run the narrow patient-absence restore smoke:
1. create one patient absence in DAILY_INPUT and apply,
2. verify MASTER treatment remains visible/red/struck,
3. verify MASTER status/reason is red but not struck,
4. verify ledger row appears under THERAPIST_DAILY,
5. restore that one row,
6. verify source DAILY_INPUT row clears,
7. verify THERAPIST_DAILY and MASTER return to recurring/base state,
8. verify MASTER status returns to base (`παρών` for this specimen),
9. verify ledger row remains as `Επαναφέρθηκε`,
10. workbook remains open in-session.

Do not run CloseOperationalDay until this native restore gate is green.
