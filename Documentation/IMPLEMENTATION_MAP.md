# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Daily effective state: MASTER + THERAPIST_DAILY + ledger

Locked target:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

Το stress case ΜΕΝΤΗΣ είναι πλέον native green για:

- active move `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- MASTER + THERAPIST_DAILY sync,
- source trace,
- individual restore,
- readable `hh:mm` time axis,
- stale robotic-orange leak fix.

Statistics integration παραμένει open.

### 2. Legacy `ΑΝΑΝΕΩΣΗ` cleanup

V85 control inventory έδειξε ακριβώς:

- Form Control / Shape type `8`,
- name `Button 1`,
- caption `ΑΝΑΝΕΩΣΗ`,
- macro `...!RefreshAll`,
- όχι ActiveX (`OLEObjects: 0`).

Το προηγούμενο generic caption cleanup δεν ήταν αρκετό λόγω COM behavior του Form Control.

Νέο surgical fix:

- direct enumeration `ws.Buttons()`,
- delete μόνο με triple identity `Button 1` + `ΑΝΑΝΕΩΣΗ` + `RefreshAll`,
- shape fallback με τα ίδια τρία identifiers,
- κανένα `ekaLedger_*` ή toolbar control δεν στοχεύεται.

Commits:

- `e62ce5354752a61c4b10568ba184991b51b24c67`
- `5c51e8988797ce67358a6886e5e41f9bf5bb848c`

Status: **implemented, native verification pending**.

### 3. Reversible daily change ledger

Native green:

- per-row `Επαναφορά`,
- MASTER sync,
- source trace,
- styling cleanup,
- time-axis display.

Open:

- selected/all restore σε multi-change preview,
- production restore = authoritative overlay removal + recompute, όχι snapshot restore.

### 4. Ορφανοί: operational queue με τελικό outcome

Κάθε unresolved/displaced patient πρέπει να καταλήγει σε:

- `Τοποθετήθηκε` σε έγκυρο therapist/time, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Rules:

- κανένας ορφανός δεν εξαφανίζεται σιωπηλά,
- `Αναστολή θεραπείας` = daily overlay/status, όχι recurring edit,
- τροφοδοτεί daily statistics ως μη πραγματοποιηθείσα θεραπεία με explicit reason,
- displaced existing patients προηγούνται των νέων ασθενών.

### 5. Dynamic daily recomputation

Locked order:

1. recurring/base schedule για concrete date,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

Recompute idempotent, χωρίς destructive recurring edits.

### 6. Replacement operational load semantics

- authoritative load = occupied timeslots,
- shared slot με πολλούς patients = ένα occupied timeslot,
- ranking πρώτα occupied timeslots, μετά session count,
- output `load x/y | patients n`.

### 7. Outpatient recurring create/edit

Παραμένει open και επιστρέφει αμέσως μετά το current daily-state/ledger/orphan milestone.

## Τι έχει ολοκληρωθεί

- MASTER operational home + accepted V81 toolbar/header UX.
- DAILY_INPUT authoritative flow.
- Patient-centric multi-treatment scheduling.
- Edit-existing-program preload + atomic resave.
- Conflict validation + explicit override.
- Optional manual provider/name persistence.
- Retired single-treatment cleanup.
- Full regression checkpoint: **699 passed, 19 warnings**.
- Cancellation slot release support.
- Replacement load + mixed-slot gate: **31 passed in 2.04s**.
- MASTER daily overlay + restore gate: **24 passed, 5 warnings in 4.39s**.
- V82 individual restore: native green.
- V83 MASTER sync + restore: native green.
- V84 time-axis + restore: native green.
- V85 stale-orange cleanup + restore: native green.

## Τρέχον verification gate

1. Run focused legacy-button regression.
2. Build fresh preview.
3. Verify `ΑΝΑΝΕΩΣΗ` is gone.
4. Verify `ekaToolbar_*` and `ekaLedger_*` controls remain.
5. If green, close presentation cleanup milestone.
6. Implement orphan outcomes + `Αναστολή θεραπείας` statistics.
7. Test selected/all restore with multiple changes.
8. Replace snapshot restore with authoritative overlay removal + recompute.
9. Resume outpatient recurring create/edit gate.
10. Clean Python 3.14 ZipFile teardown warnings later.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint.
- Outpatients cannot receive Ρομποτικό.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηλά.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency allowed.
- `PATIENT_PLANNER` authoritative για inpatient recurring, `OUTPATIENT_SCHEDULE` για outpatient recurring.
- Daily cancellations/absences/replacements/suspensions/redistributions είναι overlays, όχι recurring edits.
- Existing displaced patients προηγούνται των νέων ασθενών.
- Orphans πρέπει να τοποθετηθούν ή να πάρουν explicit `Αναστολή θεραπείας`.
- MASTER και THERAPIST_DAILY πρέπει να συμφωνούν στο effective daily state.
- Operational capacity είναι timeslot-based.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
