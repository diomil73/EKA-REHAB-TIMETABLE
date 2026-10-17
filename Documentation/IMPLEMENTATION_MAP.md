# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Authoritative restore + recompute

Το restart-safe daily truth είναι native green από V88 έως V92.

Το restore πλέον σημαίνει authoritative replacement-fact lifecycle + πλήρες recompute, όχι cell copy-back.

### 2. Replacement ID hardening

Implemented:

- first placement keeps legacy id `daily-input:<session_id>:<date>`,
- re-placement after restore gets `:v2`,
- later cycles get `:v3`, etc.,
- full replacement history including RESTORED rows drives version allocation.

This prevents an old restore request from targeting a later change.

### 3. Selected/all authoritative restore

Implemented and native green:

- `daily_replacement_fact_batch.py` validates the whole requested set before mutation,
- missing/stale/already-restored member aborts the complete batch,
- no partial restore,
- selected restore accepts multiple exact ids,
- restore-all selects all currently ACTIVE facts,
- one batch mutation + one recompute + one statistics rewrite.

Tool:

- `Python/tools/restore_daily_replacement_facts.py`

V91 selected restore and V92 restore-all both passed fresh reconstruction and statistics equivalence.

### 4. Ορφανοί + treatment suspension

Native green through the full V87-V92 lifecycle.

Locked rule: κάθε orphan καταλήγει σε valid therapist/time placement ή `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα. Recurring/base δεν αλλάζει. Finalized day μόνο με unresolved=0.

### 5. Daily statistics

`DAILY_STATISTICS` derives from the same DailySessionState. V88-V92 prove fresh reconstruction remains consistent with persisted statistics.

### 6. Daily effective state: MASTER + THERAPIST_DAILY

Native green for active replacement and authoritative restore flows. Restore resets recurring truth, reapplies surviving ACTIVE overlays, and rebuilds operational views.

### 7. Excel ledger migration

Current active implementation.

Backend bridge:

- `restore_daily_replacement_facts.py` supports `--response` JSON for Excel/VBA,
- backend bridge gate: **30 passed in 4.31s**.

The first native per-row bridge attempt proved backend preview/recompute but exposed two UI lifecycle regressions: visible `WScript.Shell.Exec` console and `ThisWorkbook.Close` failure from the active macro. The close/reopen design is therefore retired from the primary ledger UI path.

Current in-session bridge implementation:

- base ledger VBA remains in `Python/rehab_excel/daily_change_ledger_vba.py`,
- `Python/rehab_excel/daily_restore_in_session_vba.py` transforms the obsolete close/reopen leg into the current in-session transaction path,
- `Python/tools/build_daily_restore_bridge_preview.py` installs the transformed VBA into smoke copies,
- per-row / selected / all still send the same authoritative immutable-id requests,
- Python runs hidden and blocking through `WScript.Shell.Run(..., 0, True)`, so no console window is intended,
- Python still creates and validates a separate authoritative preview,
- Excel opens that verified preview with events/screen updating disabled and synchronizes the authoritative operational sheet set into the already-open workbook,
- the sheet set is `DAILY_REPLACEMENT_FACTS`, `DAILY_CHANGE_LEDGER`, `THERAPIST_DAILY`, `MASTER_SCHEDULE`, `DAILY_STATISTICS`,
- an in-session rollback copy is created before synchronization,
- the workbook is saved without `ThisWorkbook.Close` and without automatic reopen,
- regression tests forbid the old `Shell.Exec`, close/reopen worker, and `ThisWorkbook.Close SaveChanges:=True` markers in the effective bridge.

The old VBA direct snapshot cell-copy restore path remains retired. Snapshot cells are audit evidence only.

Implementation commits for the in-session correction:

- `446f6f599a30f4a561b4be8a53967d1e02703e69`: in-session transformer/bridge added,
- `b7fdb206c2e9e3b3ed5758beb52e64d7056f5a54`: smoke builder installs effective in-session VBA,
- `8776b690f53d3b4972f87ad311da9efc863173dd`: ledger regression tests switched to effective VBA and forbid close/reopen,
- `61a2f703936ea73dd1834095cf15e9ecf671a7f4`: preview-builder tests updated,
- `762b67dd086ad496f0c9cf3ac08927f955657eea`: dedicated transformer regressions,
- `2e67d5364c67d86ca51a2f795d8522c2326da3a3`: VBA sheet-iteration/path hardening.

Pending:

- Windows source test gate for the new in-session implementation,
- rebuild of `LEDGER_ROW_SMOKE.xlsm`,
- native per-row smoke with workbook remaining open and no console window,
- only then selected/all native smoke.

### 8. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply ACTIVE persisted replacements,
6. remaining orphans resolve by placement or suspension,
7. new patients only after existing displaced patients are resolved.

Recompute remains idempotent and order-independent.

### 9. Outpatient recurring create/edit

Implemented but native verification remains open after ledger authoritative-restore migration.

### 10. Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cleanup. Not a blocker.

## Ολοκληρωμένα checkpoints

- full regression: **699 passed, 19 warnings**
- load + mixed-slot: **31 passed in 2.04s**
- MASTER overlay + restore: **24 passed, 5 warnings in 4.39s**
- suspension/statistics domain: **21 passed in 3.11s**
- orphan operational: **43 passed, 3 warnings in 5.15s**
- replacement fact focused gate: **28 passed, 3 warnings in 5.14s**
- versioned-id + restore focused gate: **36 passed in 4.74s**
- rerun persistence gate: **37 passed in 7.13s**
- restore JSON bridge/backend gate: **30 passed in 4.31s**
- pre-regression ledger source gate: **42 passed in 4.27s**
- fact-to-ledger synchronization gate: **18 passed in 4.40s**
- V87 mixed orphan flow: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**
- V90 re-placement + versioned lifecycle: **native green**
- V91 selected authoritative restore: **native green**
- V92 restore-all authoritative restore: **native green**

Python 3.14 ZipFile teardown warnings remain non-blocking technical debt.

## Τρέχον verification gate

1. Pull latest branch.
2. Run the focused in-session ledger source tests.
3. If green, rebuild only `LEDGER_ROW_SMOKE.xlsm` from trusted `RERESOLVED_V90.xlsm`.
4. Open only that copy and run one per-row restore.
5. Require: no console window, workbook stays open throughout, restored fact/ledger status is correct, MASTER/THERAPIST_DAILY agree, `DAILY_STATISTICS` is exact, and VBA remains preserved.
6. Run fresh reconstruction/inventory on that copy.
7. Only after per-row native green, rebuild/test selected two-fact restore and restore-all independently.
8. Close ledger migration and resume outpatient create/edit native gate.

## Business rules που δεν πρέπει να παλινδρομήσουν

- daily facts are overlays, never destructive recurring edits,
- operational capacity is timeslot-based,
- existing displaced patients precede new patients,
- orphan never silently disappears,
- MASTER and THERAPIST_DAILY agree on effective state,
- old restore cannot overwrite a newer change,
- bulk restore is validate-all-before-mutate,
- Excel restore controls mutate authoritative facts, not operational cells directly,
- primary ledger restore UX must not close/reopen the workbook or expose a console window,
- layout write zones remain narrow,
- ΦΘ provider remains solver-owned,
- outpatients cannot receive robotic treatment.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
