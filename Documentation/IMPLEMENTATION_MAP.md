# Implementation Map

## Σκοπός

Το αρχείο αυτό είναι ο ζωντανός χάρτης υλοποίησης του project. Συμπληρώνει το `UI_APPLICATION_ROADMAP.md` και ενημερώνεται σε κάθε ουσιαστικό milestone με τέσσερα πράγματα: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποιο είναι το verification gate**, **ποια είναι τα επόμενα βήματα**.

Τελευταία ενημέρωση: 2026-10-11

---

## Τι φτιάχνουμε τώρα

### Patient-centric multi-treatment scheduling

Στόχος είναι μία ενιαία ροή προγραμματισμού ασθενούς, χωρίς ξεχωριστό Save ανά θεραπεία.

Η φόρμα πρέπει να επιτρέπει:

- εμφάνιση όλων των θεραπειών του authoritative configuration,
- επιλογή πολλών θεραπειών για τον ίδιο ασθενή,
- χειροκίνητη συμπλήρωση ημέρας, ώρας και provider όπου απαιτείται,
- δυνατότητα ακύρωσης ήδη επιλεγμένης τιμής ώστε το πεδίο να ξαναγίνεται ελεύθερο,
- μερική χειροκίνητη συμπλήρωση και automatic completion μόνο των κενών πεδίων,
- πλήρως automatic proposal όταν ο χρήστης αφήνει scheduling πεδία κενά,
- διατήρηση manual values ως hard constraints,
- strong soft preference ώστε το Ανακλινόμενο να τοποθετείται πριν ή μετά τη ΦΘ όπου είναι εφικτό,
- μία τελική atomic αποθήκευση όλου του patient program.

### Τρέχον configuration migration: specialty providers

Το V62 diagnostic smoke test επιβεβαίωσε ότι Εργο και Λογο αποτυγχάνουν όχι λόγω πραγματικής έλλειψης χώρου αλλά επειδή τα specialty provider pools του `SETTINGS` είναι κενά.

Η ρίζα είναι ότι η υπάρχουσα `Νέος θεραπευτής` registration ροή γράφει μόνο φυσιοθεραπευτές στο `SETTINGS!A` (`THERAPISTS_FTH`). Οι reader/provider models είχαν ήδη θέσεις για Ψυχ, Εργο και Λογο, αλλά τα legacy V48-based workbooks δεν είχαν authoritative registry values εκεί.

Υλοποιήθηκε migration helper που:

- διαβάζει μόνο ήδη υπάρχοντα provider names από `PATIENT_PLANNER`,
- χρησιμοποιεί τις στήλες `Ψυχ_Ψυχολόγος`, `Εργο_Θεραπευτής`, `Λογο_Θεραπευτής`,
- κάνει case-insensitive deduplication διατηρώντας το αρχικό display spelling/order,
- συμπληρώνει μόνο κενά specialty registries στο `SETTINGS`,
- δεν αντικαθιστά υπάρχοντα configured registries,
- γράφει τα specialty registries στις στήλες I/J/K που ήδη διαβάζει το `read_settings()`,
- εκτελείται κατά το scheduling preview build πριν εγκατασταθεί/χρησιμοποιηθεί η νέα patient-program φόρμα.

Το migration είναι **implemented, pending native Excel verification**.

---

## Τι έχει ολοκληρωθεί μέχρι τώρα

### Application / workbook foundation

- MASTER_SCHEDULE ως operational home και read-only hub.
- Application shell με κρυμμένο Excel chrome στο app mode.
- MASTER toolbar και βασική πλοήγηση.
- DAILY_INPUT authoritative flow.
- THERAPIST_DAILY recurring projection και writeback υποδομή.
- Patient, therapist, student και outpatient registration vertical slices.
- Outpatient scheduling και unified base schedule compatibility.
- MASTER projection, room sorting, infectious styling και outpatient markers.
- Robotic και outpatient recurring visual rules.

### Patient scheduling architecture

- Patient-centric draft model για πολλές θεραπείες.
- Όλες οι θεραπείες προέρχονται από το workbook/configuration.
- PATIENT_PLANNER blocks για ΦΘ, Ρομποτικό, Ανακλινόμενο, Πισίνα, Εργο, Λογο, ΕΦΑ και Ψυχ.
- Provider pools για φυσιοθεραπευτές, εργοθεραπευτές, λογοθεραπευτές και ψυχολόγους.
- Recurring workload ως peak daily load, όχι άθροισμα recurrence.
- Generalized automatic solver πέρα από ΦΘ/Ρομποτικό.
- Automatic day-pattern completion από `SETTINGS` όταν οι ημέρες μένουν κενές.
- Manual day/time/provider values ως locked constraints.
- Providerless treatment support.
- Recliner/ΦΘ adjacency preference στη βαθμολόγηση.
- Μία ενεργή θεραπεία για editing, με ορατές ημέρες/ώρα/provider στη γραμμή κάθε θεραπείας.
- `Backspace`/`Delete` clear για day/time/provider ώστε μια τιμή να μπορεί να ξαναδοθεί στον solver.
- Diagnostic failure reasons για specialty/provider scheduling failures.
- Specialty-provider migration από υπάρχουσες `PATIENT_PLANNER` αναθέσεις σε κενά `SETTINGS` registries για Ψυχ/Εργο/Λογο.

### Atomic final save

Η τελική αποθήκευση του patient program έχει transaction-like συμπεριφορά:

1. validation όλων των επιλεγμένων θεραπειών,
2. resolve όλων των PATIENT_PLANNER headers πριν από οποιοδήποτε write,
3. snapshot όλων των touched cells,
4. `Application.EnableEvents = False` πριν από το πρώτο write,
5. εγγραφή όλων των assignments,
6. επαλήθευση των τιμών,
7. ένα μόνο `ThisWorkbook.Save`,
8. rollback σε περίπτωση σφάλματος,
9. αποκατάσταση του προηγούμενου event state σε κάθε path.

Αυτό αντικαθιστά το παλιό per-treatment save flow που μπορούσε να οδηγήσει σε partial writes.

### Regression coverage

Υπάρχουν focused regressions για:

- generalized automatic scheduling,
- automatic day-pattern completion,
- specialty/provider handling,
- providerless treatment handling,
- recliner adjacency preference,
- patient-program bridge behavior,
- atomic-save safety ordering και rollback contract,
- keyboard clearing για day/time/provider combo selections,
- explicit diagnostics για missing provider pools και provider capacity/occupancy,
- specialty provider extraction/deduplication από legacy planner assignments.

---

## Τρέχον verification gate

Πριν θεωρηθεί ολοκληρωμένο το patient-program milestone πρέπει να περάσουν:

1. focused pytest suite για automatic solver, diagnostics, bridge, form VBA, recurring scheduling, registration slice και specialty-provider migration,
2. build νέου preview workbook από το clean V48 base,
3. επιβεβαίωση ότι το preview build δημιουργεί/γεμίζει τα `SETTINGS` specialty registries όταν υπάρχουν legacy planner assignments,
4. native Excel smoke test Εργο/Λογο με όλα τα scheduling fields ελεύθερα,
5. επιβεβαίωση ότι Εργο/Λογο εμφανίζουν provider names στα combo boxes,
6. επιβεβαίωση ότι `Backspace`/`Delete` καθαρίζει ημέρες, ώρα και provider και αφαιρεί την τιμή από τη γραμμή σύνοψης,
7. επιβεβαίωση ότι cleared field μπορεί να ξανασυμπληρωθεί από την Αυτόματη συμπλήρωση,
8. επιβεβαίωση ότι manual values παραμένουν αμετάβλητες μετά το automatic completion,
9. επιβεβαίωση ότι η τελική αποθήκευση γράφει όλα τα selected treatments μαζί,
10. επιβεβαίωση ότι time cells εμφανίζονται ως `hh:mm`, όχι raw Excel serial,
11. επιβεβαίωση ότι το workbook δεν κλείνει/ξανανοίγει κατά την αποθήκευση.

---

## Επόμενα βήματα

### Άμεσα

1. Τρέξιμο focused tests μαζί με το νέο `test_specialty_provider_settings.py`.
2. Build του επόμενου preview από V48 ώστε να εκτελεστεί το specialty-provider migration.
3. Native Excel smoke test για Εργο/Λογο και επιβεβαίωση ότι εμφανίζονται provider names.
4. Επανέλεγχος automatic completion Εργο/Λογο με όλα τα πεδία ελεύθερα.
5. Επανέλεγχος clear → automatic refill.
6. Native Excel smoke test του πλήρους manual / mixed / automatic flow και της atomic αποθήκευσης.
7. Full regression suite όταν περάσει το focused smoke test.
8. Επέκταση της `Νέος θεραπευτής` registration ροής ώστε νέα specialty staff να γράφονται απευθείας στα αντίστοιχα SETTINGS registries και να μην χρειάζεται migration για νέα δεδομένα.

### Μετά το patient-program milestone

- refresh dependent views μετά το final save όπου χρειάζεται,
- τελικός UX έλεγχος της φόρμας και των labels,
- cleanup παλιού single-treatment scheduling code,
- ενημέρωση `UI_APPLICATION_ROADMAP.md` και `Documentation/PROJECT_STATUS.md` με verified status,
- μετάβαση στο επόμενο ανοιχτό roadmap milestone.

---

## Κανόνας ενημέρωσης αυτού του χάρτη

Σε κάθε ουσιαστική αλλαγή ενημερώνονται τουλάχιστον:

- **Τι φτιάχνουμε τώρα**
- **Τι έχει ολοκληρωθεί μέχρι τώρα**
- **Επόμενα βήματα**
- **Verification gate**, όταν αλλάζει το test/smoke status

Δεν χαρακτηρίζεται κάτι ως ολοκληρωμένο μόνο επειδή γράφτηκε ο κώδικας. Για milestone που επηρεάζει workbook/UI, ολοκλήρωση σημαίνει και επιτυχές test/smoke verification.
