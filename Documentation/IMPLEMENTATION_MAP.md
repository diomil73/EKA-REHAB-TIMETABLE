# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Daily effective state: MASTER + THERAPIST_DAILY + ledger

Locked target:

`DailySessionState -> THERAPIST_DAILY + MASTER timetable + ledger + statistics`

Το V83/V84 έχει περάσει native το βασικό integration stress test:

- ΜΕΝΤΗΣ `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- `THERAPIST_DAILY` δείχνει σωστά active move + source trace,
- `MASTER` δείχνει επίσης το effective daily assignment,
- individual `Επαναφορά` γυρίζει και τα δύο views στο `Σκαρώνης 09:15`,
- η αριστερή time axis πλέον δείχνει σωστά `hh:mm` χωρίς `####`.

Status: **MASTER + THERAPIST_DAILY + ledger + time-axis native green για το stress case**. Statistics integration παραμένει open.

Focused source gate πριν το V83: **24 passed, 5 warnings in 4.39s**. Τα warnings είναι Python 3.14 `ZipFile.__del__` teardown warnings και δεν αποτελούν functional blocker.

### 2. THERAPIST_DAILY styling / legacy cleanup

Native V84 έδειξε δύο εναπομείναντα presentation bugs:

- ο ΜΕΝΤΗΣ εμφανίζεται πορτοκαλί στο destination παρότι δεν είναι robotic,
- το legacy κουμπί `ΑΝΑΝΕΩΣΗ` παραμένει.

Root cause styling:

- rewritten rich-text cell μπορεί να κληρονομήσει stale whole-cell orange font από παλιό robotic presentation,
- non-robotic γραμμή χωρίς explicit color τότε εμφανίζεται orange.

Fix implemented:

- πριν εφαρμοστούν semantic rich-text runs, rewritten cell reset σε default black όταν δεν ζητείται whole-cell font role,
- μόνο actual `robotic_orange` runs ξαναγίνονται orange.

Legacy fix implemented:

- preview installer αφαιρεί μόνο shape με ακριβές caption `ΑΝΑΝΕΩΣΗ`,
- ledger controls δεν στοχεύονται.

Commits:

- `f4e382e0f1039cbd9582e9138a7dacde5eabd404`
- `d83b36d0d38d782df9e3cc094eb0da18498906a1`
- `1785786c05f95ac39aa8f2e17d9efb6e8e46325c`
- `9fc4e948d0065dfbad838730584d2832ef8adef7`

Status: **implemented, focused/native verification pending**.

### 3. Reversible daily change ledger

- per-row `Επαναφορά`: native green,
- source trace διατηρείται και είναι αποδεκτό UX,
- selected/all restore παραμένει open για multi-change preview,
- production restore αργότερα θα γίνει authoritative overlay removal + recompute αντί snapshot restore.

### 4. Ορφανοί: operational queue με τελικό outcome

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει σε:

- `Τοποθετήθηκε` σε έγκυρο therapist/time, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Rules:

- κανένας ορφανός δεν εξαφανίζεται σιωπηλά,
- `Αναστολή θεραπείας` είναι daily overlay/status, όχι recurring edit,
- τροφοδοτεί daily statistics ως μη πραγματοποιηθείσα θεραπεία με αυτή την αιτία,
- displaced existing patients προηγούνται των νέων ασθενών.

### 5. Dynamic daily recomputation

Locked order:

1. recurring/base schedule για concrete date,
2. cancellations/no-shows και release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

Η σειρά εισαγωγής facts δεν αλλάζει το τελικό valid state. Recomputation idempotent.

### 6. Replacement operational load semantics

- authoritative load = occupied **timeslots**,
- shared slot με πολλούς patients = ένα occupied timeslot,
- ranking πρώτα occupied timeslots, μετά session count,
- output `load x/y | patients n`,
- confirmed: Μηλιδάκης `load 5/6 | patients 6`.

### 7. Mixed inpatient/outpatient presentation

Mixed effective cell:

- δεν μπλοκάρει valid movement,
- δεν παίρνει λάθος whole-cell outpatient fill,
- μπορεί προσωρινά να μείνει neutral μέχρι per-patient styling.

### 8. Outpatient recurring create/edit

Παραμένει open και επιστρέφει αμέσως μετά το current daily-state/ledger integration milestone.

## Τι έχει ολοκληρωθεί

- MASTER operational home + accepted top toolbar/header pinning στο V81.
- DAILY_INPUT authoritative flow.
- Patient-centric multi-treatment scheduling.
- Edit-existing-program preload + atomic resave.
- Conflict validation + explicit override.
- Optional manual provider/name persistence για Ρομποτικό / Εργο / Λογο / Ψυχ.
- Retired single-treatment cleanup.
- Full regression after cleanup: **699 passed, 19 warnings**.
- Cancellation slot release source support.
- Reversible ledger source gate: **10 passed in 3.26s**.
- Replacement load + mixed-slot gate: **31 passed in 2.04s**.
- MASTER daily overlay + restore source gate: **24 passed, 5 warnings in 4.39s**.
- V82 individual restore: native green.
- V83 MASTER active-overlay sync + restore: native green.
- V84 time-axis display + restore: native green.

## Τρέχον verification gate

1. Run focused tests for stale-font reset + legacy `ΑΝΑΝΕΩΣΗ` removal.
2. Build fresh preview.
3. Verify ΜΕΝΤΗΣ at Μηλιδάκης 10:45 is not orange unless actually robotic.
4. Verify `ΑΝΑΝΕΩΣΗ` is gone and ledger buttons remain.
5. Verify individual restore still works and both views stay synchronized.
6. Implement/verify orphan outcomes + statistics.
7. Test selected/all restore for multiple changes.
8. Replace snapshot restore with authoritative overlay removal + full daily recompute.
9. Resume outpatient recurring create/edit gate.
10. Clean up Python 3.14 ZipFile teardown warnings later unless they become functional failures.

## Business rules που δεν πρέπει να παλινδρομήσουν

- ΦΘ provider required + solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ optional free text, όχι solver constraint για inpatient flow.
- Outpatients cannot receive Ρομποτικό.
- Ανακλινόμενο / Πισίνα / ΕΦΑ χωρίς provider/name.
- Nonblank manual constraints δεν αντικαθίστανται σιωπηλά.
- Standard duration 45', Ανακλινόμενο 30'.
- Exact adjacency allowed.
- `PATIENT_PLANNER` authoritative για inpatient recurring, `OUTPATIENT_SCHEDULE` για outpatient recurring.
- Daily cancellations/absences/replacements/suspensions/redistributions είναι overlays, όχι recurring edits.
- Existing displaced patients προηγούνται των νέων ασθενών.
- Orphans πρέπει να τοποθετηθούν ή να πάρουν explicit `Αναστολή θεραπείας`.
- Daily restore αφαιρεί overlay + recompute, δεν γράφει recurring schedule από UI text.
- MASTER και THERAPIST_DAILY πρέπει να συμφωνούν στο effective daily state.
- Operational capacity είναι timeslot-based.
- Mixed inpatient/outpatient presentation ambiguity δεν ακυρώνει valid schedule movement.

## Recovery rule

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: πάντα `python -m pytest`, ποτέ bare `pytest`.
