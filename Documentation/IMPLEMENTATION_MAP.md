# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Authoritative restore + recompute

Το restart-safe daily truth είναι native green από V88 έως V92.

Το restore σημαίνει authoritative replacement-fact lifecycle + πλήρες recompute, όχι cell copy-back.

### 2. Replacement ID hardening

Implemented:

- first placement keeps legacy id `daily-input:<session_id>:<date>`,
- re-placement after restore gets `:v2`,
- later cycles get `:v3`, etc.,
- full replacement history including RESTORED rows drives version allocation.

### 3. Selected/all authoritative restore

Implemented and native green for correctness. Current ownership after the performance redesign:

- `daily_replacement_fact_batch.py` owns validate-all-before-mutate semantics through `validate_batch_restore_request`,
- `Python/tools/restore_daily_replacement_facts.py` reads the authoritative fact set once, validates the requested ids, and derives the post-restore ACTIVE replacement set in memory,
- the staging workbook is no longer opened/saved by a dedicated Excel process just to flip fact status,
- `_RestoreOutputBackend` owns persistence of selected fact rows to `RESTORED` in `DAILY_REPLACEMENT_FACTS`, together with MASTER, daily patches, statistics, ledger lifecycle, refresh, and the single output save,
- missing/stale/already-restored member still aborts the whole batch before authoritative mutation,
- selected restore accepts multiple exact ids,
- restore-all selects all currently ACTIVE facts.

Performance ownership change: implementation `c7895e28e5843df30a889b55326189764c3dec0a`.

### 4. Ορφανοί + treatment suspension

Native green through the full V87-V92 lifecycle.

Locked rule: κάθε orphan καταλήγει σε valid therapist/time placement ή `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα. Recurring/base δεν αλλάζει. Finalized day μόνο με unresolved=0.

### 5. Daily statistics

`DAILY_STATISTICS` derives from the same DailySessionState. V88-V92 prove fresh reconstruction remains consistent with persisted statistics.

### 6. Daily effective state: MASTER + THERAPIST_DAILY

Native green for active replacement and authoritative restore flows. Restore resets recurring truth, reapplies surviving ACTIVE overlays, and rebuilds operational views.

`build_master_projection` now reads its workbook metadata in one read-only openpyxl session: PATIENTS source-row ids, SETTINGS room order, and MASTER header location. This replaces three separate workbook opens without changing projection semantics. Implementation `63e6920c700cd73644d93a9fe8ffd79c365105cc`.

### 7. Excel ledger migration

Correctness is closed green. Current bridge:

- base ledger VBA: `Python/rehab_excel/daily_change_ledger_vba.py`,
- in-session transformer: `Python/rehab_excel/daily_restore_in_session_vba.py`,
- smoke builder: `Python/tools/build_daily_restore_bridge_preview.py`,
- backend CLI: `Python/tools/restore_daily_replacement_facts.py`,
- Python runs hidden and blocking through `WScript.Shell.Run(..., 0, True)`,
- verified preview is applied to the already-open workbook,
- rollback copy protects the in-session apply,
- authoritative sheet set: `DAILY_REPLACEMENT_FACTS`, `DAILY_CHANGE_LEDGER`, `THERAPIST_DAILY`, `MASTER_SCHEDULE`, `DAILY_STATISTICS`,
- no primary-path `Shell.Exec`, `ThisWorkbook.Close`, or detached reopen worker.

The old VBA direct snapshot cell-copy restore path remains retired. Snapshot cells are audit evidence only.

### 8. Restore performance / profiling

Current active engineering item.

Measured evolution:

- Backend before MASTER batching: `45.70s`,
- after MASTER range-batched COM write: `31.33s`,
- direct backend profile after read reuse: `28.794s total`, `20.911s pre-Excel`, `7.883s Excel output`,
- dominant old phases: `restore_fact_batch 8.863s`, `build_master_projection 5.833s`.

Both dominant phases have now been redesigned. The profiler at `Python/tools/profile_daily_restore_backend.py` measures pre-Excel and Excel output sub-phases and is updated for in-session replacement fact persistence.

### 9. Dynamic daily recomputation

Locked order:

1. recurring/base schedule,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/orphan sessions,
5. apply ACTIVE persisted replacements,
6. remaining orphans resolve by placement or suspension,
7. new patients only after existing displaced patients are resolved.

Recompute remains idempotent and order-independent.

### 10. Outpatient recurring create/edit

Implemented but native verification remains open until restore latency reaches acceptance.

### 11. Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cleanup. Not a blocker.

## Ολοκληρωμένα checkpoints

- full regression: **699 passed, 19 warnings**
- replacement lifecycle V87-V92: **native green**
- authoritative ledger restore per-row / selected / restore-all: **native green**
- focused source gates through MASTER batch-write: **green**

Python 3.14 ZipFile teardown warnings remain non-blocking technical debt.

## Τρέχον verification gate

1. Pull latest branch.
2. Run the focused restore/profiler/MASTER/bridge tests listed in `CURRENT_HANDOFF.md`.
3. If green, run a fresh `BACKEND_PROFILE3.xlsm` from trusted `RERESOLVED_V90.xlsm`.
4. Confirm the dedicated `restore_fact_batch` phase is gone and MASTER projection metadata time is materially lower.
5. Attack the next largest measured phase immediately.
6. Preserve source hash/VBA, validate-all-before-mutate, and exact reconstruction.
7. Keep acceptance target roughly `<=15s` total, ideally around 10s.

## Business rules που δεν πρέπει να παλινδρομήσουν

- daily facts are overlays, never destructive recurring edits,
- operational capacity is timeslot-based,
- existing displaced patients precede new patients,
- orphan never silently disappears,
- MASTER and THERAPIST_DAILY agree on effective state,
- old restore cannot overwrite a newer change,
- bulk restore is validate-all-before-mutate,
- Python does not save the staging/production `.xlsm` with openpyxl,
- authoritative fact lifecycle is persisted only to the verified output through native Excel,
- Excel restore controls mutate authoritative facts, not snapshot cells directly,
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
