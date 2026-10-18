# Current Handoff

Last updated: 2026-10-18

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER projection path focused gate: `59 passed in 6.85s`
- latest direct restore backend profile: total `16.509s`, pre-Excel/verification `7.976s`, Excel output `8.533s`
- latest `build_master_projection`: `0.089s` (down from `6.142s`)
- restore performance is accepted for now and deferred until end-stage polish
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

## DEFERRED: restore performance / UX latency

Restore latency is no longer the active engineering priority. End-user restore time around the low-to-mid 20-second range is accepted for the current development stage. Preserve the profiler and current timings for end-stage optimization.

## ACTIVE: outpatient recurring create/edit native verification

The recurring outpatient path is implemented but still needs a clean native create/edit gate. The UI already supports:

- creation through `frmOutpatientSchedule`,
- edit/preload through `EditOutpatientSchedule`,
- seven suggestion modes,
- therapist double-booking detection and popup handling,
- preview-only backend writes before authoritative promotion,
- stable `target_base_entry_id` for edits so the same authoritative row is updated rather than duplicated.

A one-step fresh smoke builder is available:

- `Python/tools/build_outpatient_recurring_smoke.py`
- initial implementation `7006383c242d8cb02ccf6a5e623da1372ef633d9`
- initial regression coverage `a031f864aa218f3521d7f02cadd64c1363896776`
- native build exposed a transient Windows file lock on the short-lived stage immediately after the edit-support COM session closed: `[WinError 32] ... .outpatient-form-stage.xlsm`
- fix `b6554bed0fd08471de043f3d98774a150e00c4b5`: retry only WinError 32 while handing the stage to the conflict-guard builder, up to 20 attempts × 0.25s; all other OSErrors still fail immediately
- regression `5ed68bd2fe22db21e09ce6de22829c39238d8bb7`: locks the retry contract and keeps source-mutation protections explicit

The builder never mutates the trusted source. It creates a short-lived stage, installs the schedule form + edit support, produces the final copy with the conflict guard, verifies source hash preservation, and removes the stage on success or failure.

## Exact resume point

1. Pull latest branch.
2. Run the focused outpatient source gate:
   `python -m pytest Tests/test_outpatient_schedule_form_vba.py Tests/test_outpatient_schedule_edit_vba.py Tests/test_outpatient_recurring_smoke_builder.py -q`
3. If green, rerun the same fresh smoke build:
   `python Python/tools/build_outpatient_recurring_smoke.py --source .\Excel\previews\RERESOLVED_V90.xlsm --output .\Excel\previews\OUTPATIENT_RECURRING_SMOKE.xlsm --overwrite`
4. Require the prior WinError 32 stage lock to be gone. If the build still fails with WinError 32 after retries, redesign the builder to install edit + conflict support in one Excel session rather than increasing the retry window.
5. Native create gate: open only the fresh smoke copy, run `ShowOutpatientScheduleForm`, create one valid outpatient recurring row, require backend preview verification and authoritative persistence, then confirm the new `OUTPATIENT_SCHEDULE` row reads back exactly.
6. Native edit gate: on a fresh smoke copy, select/edit that row through `EditOutpatientSchedule`, change time/days/therapist, require the same `outpatient:<row>` id and no duplicate row.
7. Native conflict gate: on a fresh smoke copy, attempt a known therapist/time recurring collision and require the conflict popup/guard rather than silent double-booking.
8. Do not manually save smoke workbooks outside the tested flow. Always use a fresh copy for each scenario.
9. Once create/edit/conflict are native green, move directly to the next scheduling/UI milestone. Do not return to restore performance tuning yet.

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
