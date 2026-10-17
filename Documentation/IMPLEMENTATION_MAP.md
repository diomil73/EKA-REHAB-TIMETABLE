# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ορφανοί: operational queue + date-specific suspension

Ενεργό milestone.

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει υποχρεωτικά σε:

- `Τοποθετήθηκε` σε έγκυρο therapist/time, ή
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Architecture locked:

- suspension = `DailySessionStatus.CANCELLED` + `SessionCancellationKind.TREATMENT_SUSPENDED`,
- recurring/base δεν αλλάζει,
- original slot απελευθερώνεται,
- reason παραμένει auditable,
- ημέρα finalized μόνο με `unresolved_sessions == 0`.

Domain foundation: **source green, 21 passed in 3.11s**.

Operational wiring τώρα implemented και περιμένει consolidated gate:

- sequential flow δεν επιτρέπει πλέον `leave unresolved`,
- candidate number = placement,
- `S` = `Αναστολή θεραπείας`,
- `Q` = abort whole preview,
- no-candidate orphan απαιτεί `S` ή `Q`,
- finalization safety gate πριν γραφτεί workbook,
- generated suspension γράφεται στο output `DAILY_INPUT` ως date-specific fact,
- `daily_input_reader.py` ξαναδιαβάζει το status ως `TREATMENT_SUSPENDED`,
- dropdown του `DAILY_INPUT` περιλαμβάνει `Αναστολή θεραπείας`,
- persistence writer είναι idempotent για patient/time/status.

### 2. Daily statistics

Domain statistics source green για:

- scheduled,
- performed,
- replaced,
- patient absent,
- department postponed,
- patient no-show,
- treatment suspended,
- unresolved.

Νέο `DAILY_STATISTICS` workbook sheet implemented. Γράφεται από το ίδιο finalized `DailySessionState` και περιλαμβάνει επίσης `finalized = ΝΑΙ/ΟΧΙ`.

Status: **implemented, consolidated source gate + native verification pending**.

### 3. Daily effective state: MASTER + THERAPIST_DAILY + ledger

Το stress case ΜΕΝΤΗΣ είναι native green για:

- active move `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- MASTER + THERAPIST_DAILY sync,
- source trace,
- individual restore,
- readable `hh:mm` time axis,
- stale robotic-orange leak fix.

### 4. Legacy `ΑΝΑΝΕΩΣΗ`

**Deferred/manual cleanup**. Δεν είναι blocker και δεν ξανανοίγει τώρα.

### 5. Reversible daily change ledger

Native green:

- per-row `Επαναφορά`,
- MASTER sync,
- source trace,
- styling cleanup,
- time-axis display.

Open:

- selected/all restore σε multi-change preview,
- production restore = authoritative overlay removal + recompute, όχι snapshot restore.

### 6. Dynamic daily recomputation

Locked order:

1. recurring/base schedule για concrete date,
2. cancellations/no-shows release slots,
3. therapist absences,
4. identify displaced/unscheduled existing patients,
5. redistribute στα net free slots,
6. resolve remaining orphans με placement ή `Αναστολή θεραπείας`,
7. μόνο μετά new-patient allocation.

Recompute idempotent και order-independent, χωρίς destructive recurring edits.

### 7. Replacement operational load semantics

- authoritative load = occupied timeslots,
- shared slot με πολλούς patients = ένα occupied timeslot,
- ranking πρώτα occupied timeslots, μετά session count,
- output `load x/y | patients n`.

### 8. Outpatient recurring create/edit

Παραμένει open και επιστρέφει μετά το current orphan/daily-state milestone.

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
- Suspension/statistics domain slice: **source green, 21 passed in 3.11s**.

## Τρέχον verification gate

1. Run consolidated source gate for orphan action flow, suspension persistence, DAILY_INPUT dropdown and DAILY_STATISTICS.
2. Native mixed orphan scenario: τουλάχιστον μία placement και μία suspension.
3. Verify output reread: suspension fact recognized, unresolved=0, finalized=ΝΑΙ, source unchanged, VBA preserved.
4. Test selected/all restore with multiple changes.
5. Replace snapshot restore with authoritative overlay removal + recompute.
6. Resume outpatient recurring create/edit gate.
7. Clean Python 3.14 ZipFile teardown warnings later.

Focused Windows gate:

```powershell
python -m pytest `
  Tests/test_daily_input_cancellations.py `
  Tests/test_daily_input_sheet.py `
  Tests/test_apply_daily_input_replacements_resolution.py `
  Tests/test_daily_operational_fact_writer.py `
  Tests/test_daily_statistics_writer.py `
  Tests/test_orphan_treatment_suspension.py `
  Tests/test_daily_treatment_statistics.py `
  Tests/test_daily_state.py `
  Tests/test_therapist_absence_queue.py `
  Tests/test_cancellation_releases_replacement_capacity.py -q
```

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
