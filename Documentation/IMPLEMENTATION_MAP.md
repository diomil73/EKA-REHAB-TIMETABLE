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
- χειροκίνητη συμπλήρωση όσων scheduling στοιχείων γνωρίζει ο χρήστης,
- δυνατότητα ακύρωσης ήδη επιλεγμένης τιμής ώστε το πεδίο να ξαναγίνεται ελεύθερο,
- μερική χειροκίνητη συμπλήρωση και automatic completion μόνο των κενών πεδίων,
- πλήρως automatic proposal όταν scheduling πεδία μένουν κενά,
- διατήρηση manual values ως hard constraints,
- strong soft preference ώστε το Ανακλινόμενο να τοποθετείται πριν ή μετά τη ΦΘ όπου είναι εφικτό,
- μία τελική atomic αποθήκευση όλου του patient program.

### Provider rule του patient-program flow

Η τρέχουσα επιχειρησιακή απόφαση είναι:

- **ΦΘ:** απαιτεί θεραπευτή/provider μέσα στη φόρμα προγραμματισμού.
- **Ρομποτικό, Εργο, Λογο, Ψυχ:** ο scheduler ενδιαφέρεται μόνο για **ημέρες + ώρα**. Provider/name δεν είναι προϋπόθεση για proposal ή final save.
- **Ανακλινόμενο, Πισίνα, ΕΦΑ:** επίσης day/time scheduling χωρίς provider.
- Τα ονόματα προσωπικού για Ρομποτικό/Εργο/Λογο/Ψυχ μπορούν να συμπληρώνονται αργότερα χειροκίνητα από τις υπάρχουσες λίστες/λειτουργίες.

Η προηγούμενη ιδέα για automatic migration specialty-provider names στο `SETTINGS` αφαιρέθηκε από το patient scheduling flow επειδή δεν χρειάζεται για τον πραγματικό στόχο του timeslot scheduling.

> **Σημείωση για μελλοντική επέκταση:** αν αργότερα χρειαστεί, μπορεί να προστεθεί optional provider-aware scheduling για Ρομποτικό/Εργο/Λογο/Ψυχ. Αυτό πρέπει να γίνει ως πρόσθετη δυνατότητα και όχι ως υποχρεωτικό constraint που μπλοκάρει ημέρες/ώρα.

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
- Recurring workload ως peak daily load, όχι άθροισμα recurrence.
- Generalized automatic solver πέρα από ΦΘ/Ρομποτικό.
- Automatic day-pattern completion από `SETTINGS` όταν οι ημέρες μένουν κενές.
- Manual day/time values ως locked constraints.
- Provider constraint μόνο στη ΦΘ.
- Ρομποτικό/Εργο/Λογο/Ψυχ προγραμματίζονται χωρίς υποχρεωτικό provider.
- Providerless treatment support.
- Recliner/ΦΘ adjacency preference στη βαθμολόγηση.
- Μία ενεργή θεραπεία για editing, με ορατές ημέρες/ώρα/provider στη γραμμή κάθε θεραπείας.
- `Backspace`/`Delete` clear ώστε μια τιμή να μπορεί να ξαναδοθεί στον solver.
- Diagnostic failure reasons για πραγματικές scheduling συγκρούσεις αντί για ψευδή provider failures.
- Focused provider-rule regression gate verified: **46 passed in 2.54s**.

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
- ΦΘ provider constraint,
- provider-optional Ρομποτικό/Εργο/Λογο/Ψυχ,
- providerless treatment handling,
- recliner adjacency preference,
- patient-program bridge behavior,
- atomic-save safety ordering και rollback contract,
- keyboard clearing για scheduling combo selections,
- explicit patient time-collision diagnostics.

---

## Τρέχον verification gate

Το focused pytest gate μετά την αλλαγή provider rule έχει περάσει: **46 passed in 2.54s**.

Πριν θεωρηθεί ολοκληρωμένο το patient-program milestone απομένουν:

1. build νέου preview workbook από το clean V48 base,
2. native Excel smoke test με ΦΘ + Ρομποτικό + Εργο + Λογο + λοιπές θεραπείες,
3. επιβεβαίωση ότι μόνο η ΦΘ εμφανίζει/απαιτεί provider στο scheduling flow,
4. επιβεβαίωση ότι Ρομποτικό/Εργο/Λογο/Ψυχ βρίσκουν proposal ημέρας/ώρας χωρίς provider names,
5. επιβεβαίωση ότι `Backspace`/`Delete` καθαρίζει επιλογές και το cleared field μπορεί να ξανασυμπληρωθεί από το automatic,
6. επιβεβαίωση ότι manual values παραμένουν αμετάβλητες μετά το automatic completion,
7. επιβεβαίωση ότι η τελική αποθήκευση γράφει όλα τα selected treatments μαζί,
8. επιβεβαίωση ότι time cells εμφανίζονται ως `hh:mm`, όχι raw Excel serial,
9. επιβεβαίωση ότι το workbook δεν κλείνει/ξανανοίγει κατά την αποθήκευση.

---

## Επόμενα βήματα

### Άμεσα

1. Build του επόμενου preview από V48.
2. Native Excel smoke test με Εργο/Λογο/Ρομποτικό χωρίς provider names.
3. Επανέλεγχος clear → automatic refill.
4. Native Excel smoke test του πλήρους manual / mixed / automatic flow και της atomic αποθήκευσης.
5. Full regression suite όταν περάσει το focused smoke test.

### Μελλοντική optional επέκταση

- Αν προκύψει πραγματική ανάγκη, προσθήκη optional provider-aware scheduling για Ρομποτικό/Εργο/Λογο/Ψυχ.
- Η επέκταση δεν πρέπει να κάνει provider υποχρεωτικό ούτε να εμποδίζει την εύρεση timeslot όταν το όνομα μένει κενό.

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