# Current Handoff

Last updated: 2026-10-25

## Read order

Start with `Documentation/ONBOARDING.md`. This file is the exact resume point.

## Branch / trusted state

- Branch: `feature/outpatient-scheduling`
- Baseline `Rehab_Center_System_v27_1.xlsm` remains immutable.
- Replacement-quality milestone is CLOSED end-to-end.
- Restore correctness is closed; restore performance remains deferred.
- Statistics/history milestone is active.

## Latest confirmed checkpoints

- full regression checkpoint: `699 passed, 19 warnings`
- OOXML MASTER gate: `59 passed in 6.85s`
- latest restore profile: `16.509s`
- double-booking focused gate: `13 passed in 1.63s`
- global replacement/reclining/service-policy gate: `38 passed in 5.30s`
- student production wiring gate: `39 passed in 2.35s`
- student layout/projection gate: `19 passed in 2.35s`
- unresolved/suspension gate: `24 passed in 3.74s`
- joined replacement decision gate: `21 passed in 3.95s`
- native student replacement smoke: VERIFIED
- native explicit REJECT branch: VERIFIED, 1 persisted `Αναστολή θεραπείας`, unresolved 0, application contract OK (54 components)
- native explicit ACCEPT branch: VERIFIED, ACTIVE fact `daily-global:planner:92:ΦΘ@2026-10-26:2026-10-26`, `Βασιλειάδης -> Μηλιδάκης 08:30`
- native finalized-day snapshot smoke: VERIFIED on `2026-10-26`, `CURRENT_FINAL`, 79 final sessions, 1 treatment suspension, unresolved 0, source unchanged, application contract `54 -> 55` components
- native snapshot + therapist/day history smoke: VERIFIED on `2026-10-26`, 16 therapist/day rows, unresolved 0, source unchanged, application contract `54 -> 56` components

## Closed replacement decision chain

Production order is locked as:

`recurring/base -> cancellations/no-shows -> therapist absences -> surviving ACTIVE replacements -> global solve -> double-booking review -> explicit ACCEPT/REJECT -> remaining suspension -> final recompute`

Double-booking remains explicit-only. Final operational state requires `unresolved_sessions == 0`.

## Student presentation/lifecycle contract

- active registered students display by real name;
- `φοιτ N` labels are physical placeholders only;
- provider kind remains `STUDENT`;
- placement end removes active real-name presentation while history retains real StudentID/name;
- physical-slot overflow safety-stops.

## Therapist workload contract

MASTER badges use `Όνομα n/6`, where `n` is occupied timeslots, not patient lines:

- 0-4 green
- 5 yellow/orange
- 6+ red

History preserves occupied timeslots and active sessions as separate metrics. A legal shared slot with two patients is one occupied timeslot and two active sessions.

## Day-finalization/history implementation state

Finalization is explicit via `Κλείσιμο ημέρας`, never automatic.

Implemented and native-verified:

- immutable `FinalizedDaySnapshot`;
- strict unresolved=0 finalization gate;
- `FINALIZED_DAY_SNAPSHOTS` hidden/protected native COM store;
- one `CURRENT_FINAL` per date;
- exact semantic read-back before Save;
- official therapist-registry wiring for DAILY_INPUT;
- source workbook hash verification;
- therapist/day historical workload derivation;
- `THERAPIST_DAY_HISTORY` hidden/protected COM store;
- 16 therapist/day rows persisted and reread on the verified 2026-10-26 fixture;
- application-contract verification after history persistence.

Implemented, awaiting focused/native gate:

- immutable append-only `DayHistoryEvent` model;
- `DAY_FINALIZED` event type linked to `related_snapshot_id`;
- `DAY_HISTORY_EVENTS` hidden/protected COM store;
- duplicate EventId rejection and semantic COM read-back;
- `Python/tools/close_day.py` verified-preview action contract;
- `modCloseDayAction` VBA module with explicit confirmation;
- authoritative commit worker handoff and reopen target `MASTER_SCHEDULE`.

The close-day VBA module is not yet installed into the MASTER toolbar. Installation is deliberately held until the new backend/event focused gate and native `close_day.py` smoke are green.

Current capacity history uses the locked operational 6-timeslot therapist baseline until a per-therapist capacity registry is introduced.

## Patient-history requirements already locked

Preserve/derive:

- first operational entry date;
- end/archive date;
- stay duration;
- treatment/service days;
- scheduled/performed/missed sessions and structured reasons;
- expected working days excluding weekends/holidays;
- treatment counts and timeline;
- totals/trends by period;
- manual append-only 15-day evaluation milestones with `Από`, `Έως`, and green/grey/red outcome;
- aggregate patient reporting for all patients and by originating clinic.

Patient lifecycle dates are manually correctable for migration of existing patients. MASTER should later show a compact start-date label such as `Έν. 12/09/26` near the responsible doctor where layout permits.

## Exact resume point

**Do not open another replacement vertical.** The replacement-quality milestone is closed unless a concrete regression appears.

Immediate next gate:

1. run focused day-event/store/close-action tests;
2. run `Python/tools/close_day.py` against `Excel/previews/DOUBLE_BOOKING_REJECTED.xlsm` into a fresh preview;
3. require snapshot + 16 therapist/day rows + exactly one `DAY_FINALIZED` event + unresolved 0 + source unchanged;
4. if green, install `modCloseDayAction` and a `Κλείσιμο ημέρας` button into the application preview/MASTER toolbar and native-smoke the actual Excel action;
5. then begin patient/day finalized outcome history so therapist, department and patient reports reconcile from the same finalized truth.

## Safety reminders

- baseline immutable;
- daily facts are overlays, never destructive recurring edits;
- native Excel/COM for real `.xlsm` mutation;
- no production/staging `.xlsm` save with openpyxl;
- fresh preview/smoke copies;
- verify source hashes;
- Windows tests use `python -m pytest`;
- no manual Save outside a tested scenario;
- later corrections append/version history instead of rewriting old finals.

## Documentation rule

Every agreed design/product/future-feature idea must be recorded in `Documentation/FUTURE_FEATURES.md` or the appropriate roadmap/decision document when agreed.
