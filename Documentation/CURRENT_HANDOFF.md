# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Trusted daily source for restore smoke copies: `Excel/previews/RERESOLVED_V90.xlsm`
- V81 MASTER toolbar/header UX remains native green.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- replacement fact focused gate: `28 passed, 3 warnings in 5.14s`
- versioned-id + single/batch restore focused gate: `36 passed in 4.74s`
- rerun gate after V90 persistence fix: `37 passed in 7.13s`
- restore JSON bridge/backend gate: `30 passed in 4.31s`
- ledger authoritative bridge source gate: `42 passed in 4.27s`
- initial restore bridge preview-builder gate: `16 passed in 2.94s`
- fact-to-ledger synchronization gate: `18 passed in 4.40s`
- parity rebuild: row/selected/all smoke copies all reported 5 ACTIVE ledger rows, fact sync=True, source unchanged=True, VBA preserved=True
- first native per-row close/reopen design: failed at workbook lifecycle with visible console + `Method 'Close' of object '_Workbook' failed`
- in-session per-row restore: native green
- in-session selected two-row restore: native correctness green, no visible console, no close/reopen; user reported noticeable latency
- replacement lifecycle history V87-V92: native green

## CLOSED: authoritative daily truth and replacement lifecycle

Concrete daily truth is reconstructable from:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

Replacement facts preserve immutable history. First placement keeps the legacy id, re-placement after restore gets `:v2`, later cycles get `:v3`, etc. Restore never rewrites a newer lifecycle fact.

Atomic selected/all backend restore is implemented and native green independently through V91/V92: validate the whole requested set before mutation, no partial restore, one fact mutation batch, one daily-state recompute, one statistics rewrite, source/VBA preservation.

## ACTIVE: Excel ledger authoritative bridge

### Projection parity fix

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a projection. The smoke-copy builder rebuilds the ledger from durable facts and preserves matching audit/snapshot extras by immutable ReplacementId. This fixed the earlier stale-ledger case of 1 ACTIVE row versus 5 authoritative ACTIVE facts.

### In-session redesign

The first real Excel button smoke proved that close/reopen was the wrong primary UX. The bridge was redesigned to follow the earlier registration solution:

- Python still performs authoritative restore validation and full recompute into a verified preview.
- backend launch is hidden and blocking, with no visible console window.
- the already-open workbook stays open.
- a rollback copy is created before applying the verified result.
- verified operational sheets are synchronized into the open workbook and saved in-session.
- primary ledger actions do not use `ThisWorkbook.Close` or the detached reopen worker.
- the success message is user-facing only: `Η επαναφορά ολοκληρώθηκε επιτυχώς.`

Key implementation:

- `Python/rehab_excel/daily_restore_in_session_vba.py`
- `Python/tools/build_daily_restore_bridge_preview.py`
- `Tests/test_daily_restore_in_session_vba.py`

### NATIVE GREEN: per-row restore

Fresh `LEDGER_ROW_SMOKE.xlsm` was built from trusted V90 with:

- Active ledger rows: 5
- Ledger synchronized from authoritative facts: True
- Source unchanged: True
- VBA preserved: True
- Restore execution mode: hidden Python + in-session Excel apply

One ACTIVE ledger row was restored from Excel. The workbook stayed open and the action completed successfully.

Post-action fact inventory:

- total historical facts: 6
- RESTORED: 2
- ACTIVE: 4
- the newly restored fact was `daily-input:planner:33:ΦΘ@2026-10-19:2026-10-19`
- the unrelated newer `daily-input:planner:57:ΦΘ@2026-10-19:2026-10-19:v2` remained ACTIVE

Fresh reconstruction then passed exactly:

- persisted replacements reread: 4
- treatment suspensions reread: 2
- reconstructed unresolved: 1
- reconstructed finalized: ΟΧΙ
- stored `DAILY_STATISTICS` matches reconstructed daily state: True

This closes the **per-row native restore gate**.

### NATIVE GREEN: selected restore correctness

The independent selected smoke restored exactly two ACTIVE rows through `Επαναφορά επιλεγμένων`. The workbook stayed open, no console appeared, and the user confirmed that the reconstruction and replacement inventory were fully green. This closes the selected correctness gate.

A noticeable execution delay remains a performance/UX item. Do not weaken the authoritative preview/recompute/rollback guarantees merely to hide the delay; profile before optimizing.

### Greek-only user-facing reason labels

User-facing `Αιτία` values must not expose technical English labels. The current bridge now translates the known legacy/internal replacement reasons `DAILY_INPUT therapist absence` and `temporary replacement` to the agreed Greek label `Απουσία θεραπευτή` when rendering the ledger in `THERAPIST_DAILY`. Internal immutable IDs remain technical and unchanged.

Implementation commits:

- `9bfa145175a0a7abc876084e2c767e2d1e570a7a` localizes ledger reason rendering
- `74eb411e2e3eda455e5c920605fac013dbfd368a` adds regression coverage
- `9cd65165c9d14dbc52bb215e777c7acb6c2addb6` records the durable UI-language rule

## Exact resume point

1. Pull the latest branch and run the focused bridge tests after the Greek-only reason patch.
2. Rebuild `LEDGER_ALL_SMOKE.xlsm` from trusted V90 so it contains the newest bridge.
3. Open only `LEDGER_ALL_SMOKE.xlsm`, confirm the visible `Αιτία` column shows Greek business wording only, then run `Επαναφορά όλων`.
4. Restore-all expectation: 0 ACTIVE replacements, 5 unresolved, finalized=ΟΧΙ, 2 treatment suspensions, exact stored statistics.
5. Require no visible console, no workbook close/reopen, correct ledger statuses, MASTER/THERAPIST_DAILY agreement, source/VBA safety.
6. After restore-all is native green, close the ledger migration correctness milestone.
7. Keep the selected-restore latency as an explicit performance/UX item to profile next; likely hotspots are the full Python preview/recompute cycle and in-session used-range synchronization.
8. Then resume outpatient recurring create/edit native gate.

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

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
