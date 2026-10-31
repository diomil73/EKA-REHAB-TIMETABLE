# Recovery pre-I green gate — 2026-10-31

Focused recovery gate result: **66 passed, 1 warning in 3.26s**.

Validated together before native I build:

- read-only daily-change restore preparation;
- atomic multi-change restore preparation for one / selected / all changes;
- native DAILY_CHANGE_LEDGER module contract;
- DAILY_INPUT CHANGE records in the same authoritative payload as operational patches;
- MASTER patient-absence strike-through + readable non-struck reason/status;
- application contract with 24 required VBA components and DAILY_CHANGE_LEDGER sheet;
- in-session production-action safety;
- unified builder and workload-module canonical build expectations.

The remaining warning is the known Python 3.14 ZipFile cleanup warning and is not treated as a functional failure.

## Native I gate

Build only a fresh artifact named `UNIFIED_CLOSE_DAY_RECOVERY_20261031_I.xlsm` with no overwrite. Production source remains immutable.

First native scenario:

1. mark the same patient absent in DAILY_INPUT;
2. apply DAILY_INPUT;
3. MASTER keeps schedule visible with red strike-through and readable non-struck reason;
4. THERAPIST_DAILY shows the resulting change in the ledger section below the schedule;
5. ledger row shows patient, From, To/absence state, reason, and Active change status;
6. use row-level `Επαναφορά`;
7. patient returns to normal THERAPIST_DAILY and MASTER state;
8. the source DAILY_INPUT row is cleared;
9. ledger row remains visible as historical `Επαναφέρθηκε`;
10. same workbook stays open; no close/reopen or preview-sync path.

If row restore passes, then test selected restore and restore-all with multiple supported patient/cancellation changes. Therapist-absence restore remains intentionally fail-closed until coordinated related replacement-fact rollback is implemented and tested.

Do not run CloseOperationalDay on I until the daily-change restore cycle passes.
