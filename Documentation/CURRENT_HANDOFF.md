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
- direct backend profile after fact-session consolidation: total `22.924s`, pre-Excel/verification `12.978s`, Excel output `9.946s`
- latest dominant pre-Excel phases: `build_master_projection 6.142s`, `pre_excel_unattributed 2.581s`, `read_target_date 1.121s`, `read_daily_input 0.855s`, `read_base_schedule 0.724s`
- `restore_fact_batch 8.863s` has been eliminated from the pipeline
- first OOXML MASTER-projection gate after package-probe migration: `5 failed, 51 passed in 5.06s`
- failure root cause: openpyxl-generated relationship targets may be package-absolute, e.g. `/xl/worksheets/sheet1.xml`, while `zipfile` archive members have no leading slash
- fix `0057aa38b9fc90b66a4bd5bef543a82edf5262b7`: normalize package-absolute and relative OPC relationship targets in `WorkbookPackageProbe._join_part`
- regression `65dc8827dfa0d3f047d1823b68682f6232b2b395`: direct coverage for absolute, relative, and parent-relative relationship targets
- **post-fix focused gate: `59 passed in 6.85s`**
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

The latest profile confirms the previous fact-lifecycle change worked: backend total fell to `22.924s` and the old `restore_fact_batch 8.863s` phase disappeared. The dominant remaining pre-Excel cost is `build_master_projection 6.142s`.

MASTER projection metadata has been moved from full openpyxl workbook loads to the read-only OOXML package probe. The first Windows gate exposed an OPC-path compatibility bug rather than a projection-rule regression: relationship targets beginning with `/` were passed directly to `zipfile`, which expects archive member names such as `xl/worksheets/sheet1.xml`. The package probe now resolves both package-absolute and relative relationship targets to normalized ZIP member names before reading them. The complete focused gate is now green at `59 passed in 6.85s`.

OOXML projection implementation `82abf70349d52aaef627f7a83f5f2387d0b6f74a`, initial projection regression `5cf0fdfee7ad0c7d458b9b0e85c5e364c5a6271d`, OPC target fix `0057aa38b9fc90b66a4bd5bef543a82edf5262b7`, target-resolution regression `65dc8827dfa0d3f047d1823b68682f6232b2b395`.

Safety semantics are unchanged. The package path only reads cached/stored metadata needed for projection row mapping; it does not evaluate formulas or save the workbook.

## Exact resume point

1. Run the backend profiler with a fresh output:
   `python Python/tools/profile_daily_restore_backend.py --input .\Excel\previews\RERESOLVED_V90.xlsm --output .\Excel\previews\BACKEND_PROFILE4.xlsm --all-active --overwrite`
2. Primary expectation: `build_master_projection` should drop sharply from `6.142s`.
3. If backend remains above target, attack the next largest measured phase immediately. The current likely target is `pre_excel_unattributed ~2.58s`, then Excel process startup/open/save overhead.
4. Re-run fresh reconstruction after performance changes.
5. Keep the acceptance target at roughly `<=15s` total, ideally around 10s.
6. Once restore latency is acceptable, resume outpatient recurring create/edit native gate.

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
