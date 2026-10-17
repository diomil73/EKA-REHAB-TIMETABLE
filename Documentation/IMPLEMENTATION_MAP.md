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

- `restore_daily_replacement_facts.py` now supports `--response` JSON for Excel/VBA,
- backend bridge gate: **30 passed in 4.31s**.

VBA bridge in `Python/rehab_excel/daily_change_ledger_vba.py` now does:

- per-row: one exact immutable ReplacementId,
- selected: one atomic request with repeated exact ids,
- restore-all: one `--all-active`,
- one backend call per action,
- safe preview response + source SHA,
- handoff to `authoritative_commit_worker_cli.py`,
- workbook close, verified authoritative commit, automatic reopen on `THERAPIST_DAILY`.

The old VBA `Cells(...).Copy Destination:=...` restore path has been removed from the ledger module. Snapshot cells remain audit evidence only.

Pending:

- source gate after VBA wiring,
- Windows-native smoke for row / selected / all,
- final retirement marker once native bridge is green.

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
- V87 mixed orphan flow: **native green**
- V88 full restart reconstruction: **native green**
- V89 single authoritative restore + reconstruction: **native green**
- V90 re-placement + versioned lifecycle: **native green**
- V91 selected authoritative restore: **native green**
- V92 restore-all authoritative restore: **native green**

Python 3.14 ZipFile teardown warnings remain non-blocking technical debt.

## Τρέχον verification gate

1. Pull latest branch.
2. Run ledger bridge source tests plus authoritative restore backend tests.
3. Build fresh native preview with the updated ledger VBA module.
4. Test per-row restore from Excel.
5. Test selected two-fact restore from Excel.
6. Test restore-all from Excel on an independent trusted copy.
7. Verify fact lifecycle, MASTER/THERAPIST_DAILY, ledger status, `DAILY_STATISTICS`, source unchanged, VBA preserved.
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
