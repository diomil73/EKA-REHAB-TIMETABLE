# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- focused gate after MASTER batch-write patch: `35 passed in 3.67s`
- diagnostic restore-all before MASTER batch write: Backend `45.70s`
- diagnostic restore-all after MASTER batch write: Backend `31.33s`, rollback `0.15s`, preview open `0.88s`, sheet sync `0.19s`, final save `1.23s`
- direct backend profile after read-reuse work: total `28.794s`, pre-Excel/verification `20.911s`, Excel output `7.883s`
- dominant measured pre-Excel phases: `restore_fact_batch 8.863s`, `build_master_projection 5.833s`, `pre_excel_unattributed 2.667s`
- Excel output detail from that run: dispatch `1.528s`, open `1.268s`, MASTER `1.612s`, daily patches `0.540s`, statistics `0.271s`, ledger rows `0.022s`, refresh `0.037s`, save `1.666s`, close `0.202s`, quit `0.048s`
- parity rebuild: row/selected/all smoke copies authoritative and VBA-preserving
- per-row / selected / restore-all in-session correctness: native green
- replacement lifecycle history V87-V92: native green

## CLOSED: authoritative daily truth and replacement lifecycle

Concrete daily truth remains reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Replacement facts preserve immutable history. First placement keeps the legacy id, re-placement after restore gets `:v2`, later cycles get `:v3`, etc. Restore never rewrites a newer lifecycle fact.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a projection. The primary Excel restore path keeps the workbook open, launches Python hidden and blocking, creates rollback protection, imports the verified operational state in-session, and saves without close/reopen or visible console.

Native correctness is green for per-row, selected multi-row, and restore-all controls. User-facing `Αιτία` remains Greek-only for known replacement causes.

## ACTIVE: restore performance / UX

The latest profiler isolated two dominant costs: a separate Excel session used only to flip replacement-fact statuses (`8.863s`) and three repeated openpyxl workbook opens inside MASTER projection (`5.833s`). Both are now redesigned.

New performance work on the branch:

- `c7895e28e5843df30a889b55326189764c3dec0a`: selected/all restore no longer mutates the staging workbook through a separate Excel process. The batch is validated in memory, effective ACTIVE replacements are derived by excluding the requested ids, and authoritative `DAILY_REPLACEMENT_FACTS` statuses are changed to `RESTORED` inside the already-open `_RestoreOutputBackend` workbook before its single save.
- `63e6920c700cd73644d93a9fe8ffd79c365105cc`: `build_master_projection` now gathers PATIENTS row ids, SETTINGS room order, and MASTER header position in one read-only openpyxl workbook session instead of three separate workbook loads.
- `4d928b884e81264620b9c666016a77f23aa0e96e`: profiler updated for the new lifecycle and now times `read_replacement_fact_rows`, validation/derivation, and in-session `replacement_fact_rows` persistence.
- regression coverage: `140126a6a3f09ccaa9fe4b29f8a09547e8408119`, `9a6e7243c7a02d3191b5feedbf59de4d45addfb2`, `44c8e6556849fc4a091179a43dae7962e52cf5b5`.

These changes intentionally preserve validate-all-before-mutate semantics: no authoritative fact cell is changed until the requested set has been fully validated. The stage file remains read-only from Python's perspective; only the verified output copy is mutated through native Excel.

## Exact resume point

1. Pull latest branch.
2. Run:
   `python -m pytest Tests/test_authoritative_bulk_daily_restore.py Tests/test_profile_daily_restore_backend.py Tests/test_master_projection.py Tests/test_master_daily_overlay.py Tests/test_outpatient_schedule_source.py Tests/test_daily_restore_in_session_vba.py Tests/test_daily_change_ledger_vba.py Tests/test_daily_restore_bridge_preview.py -q`
3. If green, rerun the backend profiler with a fresh output:
   `python Python/tools/profile_daily_restore_backend.py --input .\Excel\previews\RERESOLVED_V90.xlsm --output .\Excel\previews\BACKEND_PROFILE3.xlsm --all-active --overwrite`
4. Expected improvement: `restore_fact_batch` disappears entirely; `build_master_projection` should fall materially because metadata now uses one workbook open.
5. If total backend remains above target, attack the next largest measured phase immediately. The current likely next target is the `apply_write_plan_total - excel_output_total` gap / output-copy verification overhead, then remaining parser loads.
6. Re-run fresh reconstruction after performance changes.
7. Keep the acceptance target at roughly `<=15s` total, ideally around 10s.
8. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.

## Locked orphan rule

Every orphan for the concrete date resolves as exactly one of:

- valid therapist/time placement, or
- `Αναστολή θεραπείας` for that date only.

Suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`. Recurring/base remains untouched. A finalized day requires unresolved=0.

## Locked recomputation order

1. recurring/base schedule
2. cancellations/no-shows release capacity
3. therapist absences
4. identify displaced/orphan sessions
5. apply surviving ACTIVE replacements
6. resolve remaining orphans by placement or suspension
7. only then expose capacity to new patients

Operational capacity is timeslot-based. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Documentation discipline

Meaningful changes, regressions, native results, decisions, needs, planned additions/removals/fixes must be recorded in repository documentation during the same work cycle. Chat history is not the sole source of truth.

## Legacy `ΑΝΑΝΕΩΣΗ`

Deferred/manual cosmetic cleanup. Not a blocker.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
