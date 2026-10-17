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
- latest Excel output detail: dispatch `2.133s`, open `1.849s`, MASTER `1.794s`, daily patches `0.498s`, statistics `0.277s`, replacement fact rows `0.155s`, ledger rows `0.174s`, refresh `0.129s`, save `2.189s`, close `0.098s`, quit `0.004s`
- `restore_fact_batch 8.863s` has been eliminated from the pipeline
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

The latest profile confirms the previous fact-lifecycle change worked: backend total fell to `22.924s` and the old `restore_fact_batch 8.863s` phase disappeared. The dominant remaining pre-Excel cost is now `build_master_projection 6.142s`.

The one-openpyxl-session MASTER metadata change did not materially improve this because loading the full macro-enabled workbook itself remains expensive. MASTER projection metadata has therefore been moved to the existing read-only OOXML package probe. PATIENTS row ids, SETTINGS room order, and MASTER header position are now read directly from the workbook ZIP/XML package without loading the full `.xlsm` through openpyxl.

New implementation `82abf70349d52aaef627f7a83f5f2387d0b6f74a`, regression `5cf0fdfee7ad0c7d458b9b0e85c5e364c5a6271d`.

Safety semantics are unchanged. The package path only reads cached/stored metadata needed for projection row mapping; it does not evaluate formulas or save the workbook.

## Exact resume point

1. Pull latest branch.
2. Run:
   `python -m pytest Tests/test_authoritative_bulk_daily_restore.py Tests/test_profile_daily_restore_backend.py Tests/test_master_projection.py Tests/test_master_daily_overlay.py Tests/test_outpatient_schedule_source.py Tests/test_daily_restore_in_session_vba.py Tests/test_daily_change_ledger_vba.py Tests/test_daily_restore_bridge_preview.py -q`
3. If green, rerun the backend profiler with a fresh output:
   `python Python/tools/profile_daily_restore_backend.py --input .\Excel\previews\RERESOLVED_V90.xlsm --output .\Excel\previews\BACKEND_PROFILE4.xlsm --all-active --overwrite`
4. Primary expectation: `build_master_projection` should drop sharply from `6.142s`.
5. If backend remains above target, attack the next largest measured phase immediately. The current likely target is `pre_excel_unattributed ~2.58s`, then Excel process startup/open/save overhead.
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
