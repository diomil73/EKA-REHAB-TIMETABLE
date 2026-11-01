# Clean MASTER hub UI contract — 2026-11-01

The user-facing MASTER_SCHEDULE is a clean, read-only operational hub.

## Normal-user contract

- No operational or administrative buttons are installed on MASTER_SCHEDULE.
- Existing routing macros may remain in VBA for use from other surfaces, but normal users do not see their controls on MASTER.
- Operational actions belong on the relevant operational sheet or menu.
- `Αποθήκευση ημέρας` belongs on the DAILY_INPUT footer, on the same row as the existing DAILY_INPUT actions/navigation.
- `Save & Exit` remains an application-shell action available from operational sheets; it is not a MASTER-only feature.

## Super-user exception

A future privileged/super-user mode may expose additional controls, including controls on MASTER, but only through an explicit role/mode gate. Normal mode must remain clean and must not rely on hidden or merely off-screen admin buttons.

## Current implementation checkpoints

- `1ef6239` — move `Αποθήκευση ημέρας` off MASTER and into DAILY_INPUT footer.
- `ed4d944` — remove all normal-user MASTER operational buttons from the canonical master-toolbar installer.
- `4bb90f8` — update MASTER toolbar regression tests for the clean-hub contract.
- `3ce8234` / `48fb1ea` — cover the day-save placement contract and fix its regression assertion.

## Native artifact status

- AC is a failed/diagnostic H3 UI artifact and must not be repaired or reused as proof.
- The next UI smoke must use a fresh uniquely named workbook after the source gate is green.
