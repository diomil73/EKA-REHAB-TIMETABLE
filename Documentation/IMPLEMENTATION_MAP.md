# Implementation Map

Τελευταία ενημέρωση: 2026-10-17

## Σκοπός

Ζωντανός χάρτης υλοποίησης. Για ακριβές resume point πρώτο αρχείο είναι το `Documentation/CURRENT_HANDOFF.md`.

## Τι φτιάχνουμε τώρα

### 1. Ορφανοί: operational queue + date-specific suspension

Αυτό είναι το ενεργό milestone.

Κάθε unresolved/displaced patient για τη συγκεκριμένη ημερομηνία πρέπει να καταλήγει υποχρεωτικά σε ένα από:

- `Τοποθετήθηκε` σε έγκυρο therapist/time,
- `Αναστολή θεραπείας` για τη συγκεκριμένη ημέρα.

Architecture locked:

- suspension δεν γίνεται νέο `DailySessionStatus`,
- χρησιμοποιεί το υπάρχον `CANCELLED` με νέο `SessionCancellationKind.TREATMENT_SUSPENDED`,
- recurring/base δεν αλλάζει,
- το original provider slot απελευθερώνεται,
- το reason παραμένει auditable.

Source implementation ήδη μπήκε:

- `TREATMENT_SUSPENDED = "treatment_suspended"`,
- νέο `rehab_core/orphan_resolution.py`,
- `create_treatment_suspension()` με reason `Αναστολή θεραπείας`,
- νέο `rehab_core/daily_statistics.py`,
- νέα focused tests για suspension, queue removal και statistics.

Status: **source green: 21 passed in 3.11s**. Operational Excel UI/action wiring και authoritative persistence είναι τώρα το αμέσως επόμενο βήμα.

### 2. Daily statistics

Το νέο statistics layer μετρά ξεχωριστά:

- scheduled,
- performed (`ACTIVE` + `REPLACED`),
- replaced,
- patient absent,
- department postponed,
- patient no-show,
- treatment suspended,
- unresolved (`THERAPIST_ABSENT`).

Unresolved orphan δεν καταγράφεται ψευδώς ως cancellation. Η ημέρα γίνεται `finalized` μόνο όταν `unresolved_sessions == 0`.

### 3. Daily effective state: MASTER + THERAPIST_DAILY + ledger

Το stress case ΜΕΝΤΗΣ είναι native green για:

- active move `Σκαρώνης 09:15 -> Μηλιδάκης 10:45`,
- MASTER + THERAPIST_DAILY sync,
- source trace,
- individual restore,
- readable `hh:mm` time axis,
- stale robotic-orange leak fix.

### 4. Legacy `ΑΝΑΝΕΩΣΗ`

Το κουμπί είναι **deferred/manual cleanup**. Δεν είναι blocker και δεν ξανανοίγει τώρα.

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

1. Wire actual orphan actions: placement or suspension.
2. Persist suspension as authoritative date-specific fact.
3. Wire statistics into workbook/reporting.
4. Native-check orphan workflow.
5. Test selected/all restore with multiple changes.
6. Replace snapshot restore with authoritative overlay removal + recompute.
7. Resume outpatient recurring create/edit gate.
8. Clean Python 3.14 ZipFile teardown warnings later.

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
