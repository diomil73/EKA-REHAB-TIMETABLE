# Project Status

Last updated: 2026-10-18

## Current stage

Branch: `feature/outpatient-scheduling`.

Active milestone: **Restore correctness is closed green and restore performance is deferred. Active work is now outpatient recurring create/edit native verification, followed by the broader scheduling/UI roadmap.**

## Latest confirmed checkpoints

- full regression: **699 passed, 19 warnings**
- package relationship normalization + OOXML projection gate: **59 passed in 6.85s**
- latest direct restore backend profile: **16.509s total**, **7.976s pre-Excel/verification**, **8.533s Excel output**
- `build_master_projection`: **0.089s**, down from **6.142s**
- V87-V92 authoritative restore/reconstruction lifecycle: **native green**
- per-row / selected / restore-all in-session correctness: **native green**

Python 3.14 `ZipFile.__del__` teardown warnings remain non-blocking technical debt.

## CLOSED: restart-safe authoritative daily state

Operational truth is reconstructable from durable facts:

`recurring/base + DAILY_INPUT + ACTIVE DAILY_REPLACEMENT_FACTS -> DailySessionState -> statistics`

V88-V92 prove restart-safe reconstruction, immutable replacement lifecycle, atomic selected/all semantics, and exact stored-statistics equivalence.

## CLOSED: Excel ledger authoritative restore correctness

`DAILY_REPLACEMENT_FACTS` is authoritative. `DAILY_CHANGE_LEDGER` is a UI/audit projection, never an independent truth source.

The normal Excel restore workflow runs the authoritative Python backend hidden and blocking, keeps the current workbook open, creates rollback protection, imports verified operational state, saves in-session, avoids visible console/close-reopen behavior, and renders Greek-only business wording in the visible `Αιτία` column.

Native correctness is green for per-row, selected, and restore-all controls.

## DEFERRED: restore performance / UX latency

The trusted backend has fallen from about `45.70s` initially to `16.509s`. Current end-user restore duration in the low-to-mid 20-second range is accepted for this development stage. Further tuning is deferred until final polish unless a regression appears.

## ACTIVE: outpatient recurring create/edit

The functional path already supports recurring outpatient creation, stable-row editing via `target_base_entry_id`, seven suggestion modes, therapist double-booking checks, and the conflict popup. The open item is native verification of the complete Excel flow.

To remove setup ambiguity, a combined smoke builder now produces one fresh `.xlsm` containing the create form, edit support, suggestions and conflict guard in one reproducible flow:

- `Python/tools/build_outpatient_recurring_smoke.py`
- implementation `7006383c242d8cb02ccf6a5e623da1372ef633d9`
- regression coverage `a031f864aa218f3521d7f02cadd64c1363896776`

The builder preserves the trusted source hash, uses a disposable staging copy, installs edit support only on that stage, produces the final conflict-enabled smoke copy, and removes the stage on both success and failure.

## Immediate roadmap

1. Run the focused outpatient source gate including the new combined smoke builder regression.
2. Build a fresh `OUTPATIENT_RECURRING_SMOKE.xlsm` from trusted `RERESOLVED_V90.xlsm`.
3. Native create verification: valid recurring outpatient row persists exactly.
4. Native edit verification: same `outpatient:<row>` is updated with no duplicate row.
5. Native conflict verification: known therapist/time overlap surfaces the conflict guard rather than silently saving.
6. Use a fresh smoke workbook for each scenario and do not manually save outside the tested flow.
7. When these are green, continue directly to the next scheduling/UI milestone.
8. Revisit restore performance only near completion.

## Locked daily rules

1. recurring/base schedule
2. cancellations/no-shows release slots
3. therapist absences
4. identify displaced/orphan sessions
5. apply surviving ACTIVE replacements
6. unresolved orphans resolve through placement or treatment suspension
7. only then expose capacity to new patients

Operational capacity is timeslot-based. Daily facts never destructively edit recurring schedules. Existing displaced patients precede new patients. MASTER and THERAPIST_DAILY must agree.

## Safety constraints

- never write directly to baseline `.xlsm` during development
- preview workflows use copies
- preserve and verify VBA
- keep layout write zones narrow
- existing replacement facts must never be re-appended during rerun
- partial previews created before later failure must be removed
- bulk restore validates the whole set before any mutation
- Excel ledger restore mutates authoritative fact lifecycle, never old snapshot cells directly
- `DAILY_CHANGE_LEDGER` remains a projection of authoritative facts
- normal ledger UI must not close/reopen the workbook or expose a console window
- normal ledger UI must not expose internal English reason labels in `Αιτία`
- Windows tests use `python -m pytest`

## Documentation discipline

Every meaningful evolution must be recorded in repository documentation during the same work cycle. Chat history is not the sole source of truth.

## Deferred issue

Legacy `ΑΝΑΝΕΩΣΗ` Form Control remains manual/deferred cosmetic cleanup.

## Recovery order

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests
