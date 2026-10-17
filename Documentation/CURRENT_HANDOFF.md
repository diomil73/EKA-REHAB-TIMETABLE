# Current Handoff

Last updated: 2026-10-17

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V81.xlsm`
- Latest successfully generated ledger preview: `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.
- V81 MASTER toolbar/header pinning is native green.

## MASTER controls status

Native experiments V74-V79 are retired as sidebar/persistent-control attempts.

Final accepted MASTER UX:

- row 2 contains the four existing horizontal action buttons,
- row 4 contains the MASTER treatment/column headers,
- the top area is frozen through row 4,
- the application shell reapplies the pin whenever MASTER becomes active,
- no UserForm/sidebar/CommandBars mechanism is used.

Native V81 verification: green. Vertical scrolling keeps both the four MASTER buttons and row-4 treatment headers visible.

## Patient-program native verification

Confirmed green by the user:

1. edit-existing-program end to end,
2. existing program preload,
3. atomic resave without false self-conflict,
4. optional manual names persistence/reload for Ρομποτικό / Εργο / Λογο / Ψυχ.

## Retired single-treatment cleanup

Cleanup is complete and accepted.

Removed/retired source:

- `Python/rehab_excel/patient_scheduling_flow.py`
- `Python/rehab_excel/patient_scheduling_form_vba.py`
- obsolete single-treatment bridge commands and tests.

Retained current patient-program behavior and legacy component cleanup.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- cleanup/reset focused gate: `42 passed in 3.70s`
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- V81 native toolbar/header pinning gate: green
- optional manual names persistence/reload: native green
- full regression before cleanup: `722 passed, 19 warnings`
- full regression after cleanup: `699 passed, 19 warnings`
- reversible daily change ledger focused source gate: `10 passed in 3.26s`
- replacement load + mixed outpatient/inpatient regression gate: `31 passed in 2.04s`
- V82 guided replacement native build: green

## V82 native guided replacement checkpoint

The real stress case now builds successfully from V81:

- date: `2026-10-19`
- patient: `ΜΕΝΤΗΣ`
- original assignment: `Σκαρώνης 09:15`
- selected replacement: `Μηλιδάκης 10:45`
- candidate display for Μηλιδάκης: `load 5/6 | patients 6`
- source workbook unchanged: `True`
- VBA preserved: `True`
- daily change ledger installed: `yes`
- output: `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`

This confirms the timeslot-based load semantics, mixed inpatient/outpatient non-blocking rule, COM-safe ledger date write, successful preview write, and ledger installation path.

The ledger UX itself is NOT yet native-accepted. The next gate is visual/interactive inspection inside `THERAPIST_DAILY`.

## Cancellation slot release status

Cancellation propagation is already implemented and regression-covered in source:

- cancellations are passed into therapist/patient/student availability checks,
- cancelled occurrences are excluded from therapist/student operational workload,
- replacement candidate capacity reflects the released timeslot,
- `Tests/test_cancellation_releases_replacement_capacity.py` verifies workload release and exact-slot reuse.

## Locked daily-operations rule: cancellations + therapist absences

Patient cancellations for both inpatients and outpatients are operationally critical because they release same-day timeslots that can be reused for:

- patients displaced by an absent therapist,
- redistribution of active patients across therapists,
- new patients when capacity remains.

The workflow must NOT depend on one fixed chronological entry order. Therapist absences are often known earlier, while patient cancellations may arrive much later. Therefore both event types must be accepted at any time and the daily plan must be recomputable/idempotent from the latest facts.

Required evaluation model whenever the operational plan is refreshed:

1. start from the recurring/base schedule for the selected date,
2. apply all known patient cancellations/no-shows for that date and release their occupied slots,
3. apply all known therapist absences and mark those providers unavailable,
4. identify the remaining displaced/unscheduled patients,
5. redistribute those existing affected patients into the net free capacity,
6. only after that expose remaining capacity for new-patient assignment.

Important temporal behavior:

- if therapist absence is entered first, the system may create a preliminary redistribution,
- if a patient cancellation arrives later, refreshing/recalculating the same date must immediately release the slot and allow a better redistribution,
- if a cancellation is removed/corrected, recalculation must reclaim the slot and revalidate assignments,
- base recurring schedules must not be destructively edited to represent a one-day cancellation or absence.

This is now a core invariant for the future daily allocation engine and native DAILY_INPUT/absence workflows.

## Locked replacement load semantics

Operational therapist load is measured in occupied daily **timeslots**, not patient/session count.

- `active_timeslots` is the authoritative load/capacity measure.
- Multiple patients sharing the same clock slot count as multiple sessions but one occupied timeslot.
- Replacement ranking now sorts primarily by occupied timeslots, with session count as a secondary tie-breaker.
- Candidate display now distinguishes `load x/y` from `patients n`.
- Example: 6 sessions across 5 occupied slots must display operational load `5/6`, not `6`.

## Mixed inpatient/outpatient presentation rule

A mixed inpatient/outpatient effective slot must not block an otherwise valid operational preview.

- do not apply one whole-cell outpatient fill when the cell contains mixed patient types,
- do not abort the replacement preview solely because of that presentation ambiguity,
- preserve correctness of schedule movement first; mixed-cell styling can remain neutral until per-patient text styling is implemented.

## Locked UX requirement: reversible daily change ledger in THERAPIST_DAILY

Many same-day operational redistributions are temporary and should not leak into the following working day. The user requires a visible, date-scoped change ledger below the normal therapist schedules in `THERAPIST_DAILY`.

Section title:

`Αλλαγές για την dd/mm/yy`

Each effective daily move must preserve and display at least:

- patient identity/name,
- original therapist,
- original time,
- effective/replacement therapist,
- effective/replacement time,
- optional reason/source such as therapist absence, cancellation-driven redistribution, or manual move,
- change/replacement identifier and status internally.

Required restore controls:

- per-row `Επαναφορά`,
- row selection for bulk operations,
- `Επαναφορά επιλεγμένων`,
- `Επαναφορά όλων`.

Restore semantics:

- restore means remove/reverse the date-specific operational override and recompute that same day's effective plan from the surviving facts,
- it must never destructively rewrite the recurring/base schedule,
- restore must be idempotent,
- restore must validate the current effective state before undoing so that an older ledger item cannot silently overwrite a newer change,
- after a restore, capacity/conflicts/replacements for the date must be recalculated,
- bulk restore must either process a validated set safely or stop/report conflicts rather than partially corrupting the day.

Next-day behavior:

- date-specific overlays expire naturally outside their target date,
- therefore the following day starts again from the recurring/base programme plus that day's own cancellations/absences/overrides,
- no manual "undo yesterday" should be necessary merely to return to the recurring schedule.

Architecture fit already exists in core:

- `DailySessionState` preserves `original_therapist_id` and `original_time` separately from effective therapist/time,
- `ReplacementAssignment` is already a date-specific overlay tied to one target session,
- this makes the change ledger/revert workflow a natural extension of the existing daily-state model rather than a second schedule store.

### Prototype now implemented

A first native UX prototype is wired into the existing guided replacement preview:

- new module: `Python/rehab_excel/daily_change_ledger_vba.py`,
- new focused tests: `Tests/test_daily_change_ledger_vba.py`,
- `Python/tools/apply_replacement_choice.py` now installs the ledger after creating a real replacement preview,
- the ledger uses a very-hidden `DAILY_CHANGE_LEDGER` sheet for preview-only metadata and exact source-cell snapshots,
- `THERAPIST_DAILY` renders the date-scoped change section below the existing schedule,
- per-row / selected / all restore controls are installed as native Excel form controls,
- restore refuses to overwrite a cell whose current value no longer matches the expected changed state,
- restore acts only on the preview copy and restores exact source-cell snapshots; recurring source schedules remain untouched.

Focused source verification is green: **10 passed in 3.26s**.
Replacement-load / mixed-slot fixes are also green: **31 passed in 2.04s**.
V82 guided replacement build and ledger installation are native green.

This prototype is intentionally a UX/native validation slice. After native acceptance, restore will be connected to the authoritative daily overlay store and recomputation engine rather than relying on preview cell snapshots.

## Current development focus

Immediate gate: **open `DAILY_CHANGE_LEDGER_V82.xlsm` and native-smoke the reversible ledger inside `THERAPIST_DAILY`**.

Next native check:

1. open `Excel/previews/DAILY_CHANGE_LEDGER_V82.xlsm`,
2. go to `THERAPIST_DAILY`,
3. confirm ΜΕΝΤΗΣ moved from Σκαρώνης 09:15 to Μηλιδάκης 10:45,
4. scroll below the schedules and verify `Αλλαγές για την 19/10/2026`,
5. verify the row shows ΜΕΝΤΗΣ, `Σκαρώνης 09:15` and `Μηλιδάκης 10:45`,
6. click the row `Επαναφορά`,
7. confirm the exact source/destination cells return to their prior preview state,
8. confirm the ledger row remains visible as restored history,
9. confirm recurring/base schedule remains untouched.

After individual restore is native green, test selected/all restore once multiple preview changes are supported. Then replace preview-snapshot restore with authoritative overlay removal + daily recompute. The outpatient recurring create/edit work remains open and resumes after this UX spike is accepted or adjusted.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative for inpatient recurring schedules; `OUTPATIENT_SCHEDULE` authoritative for outpatient recurring schedules.
- One-day cancellations, therapist absences, replacements and manual redistributions are overlays on the daily operational state, not destructive edits of recurring source schedules.
- Existing displaced patients have allocation priority over newly added patients when released capacity is redistributed.
- Daily restore operations must remove/reverse overlays and recompute, never reconstruct the recurring schedule from UI text.
- Replacement operational load/capacity is timeslot-based, not raw session-count based.
- Mixed inpatient/outpatient cells must not abort valid schedule movement just because whole-cell colouring is ambiguous.
- Conflict override must be explicit; return-for-correction performs zero writes.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
