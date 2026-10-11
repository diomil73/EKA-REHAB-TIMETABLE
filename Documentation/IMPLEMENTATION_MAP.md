# Implementation Map

## Σκοπός

Το αρχείο αυτό είναι ο ζωντανός χάρτης υλοποίησης του project. Συμπληρώνει το `UI_APPLICATION_ROADMAP.md` και ενημερώνεται σε κάθε ουσιαστικό milestone με τρεις ενότητες: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποια είναι τα επόμενα βήματα**.

Τελευταία ενημέρωση: 2026-10-11

---

## Τι φτιάχνουμε τώρα

### Patient-centric multi-treatment scheduling

Στόχος είναι μία ενιαία ροή προγραμματισμού ασθενούς, χωρίς ξεχωριστό Save ανά θεραπεία.

Η φόρμα πρέπει να επιτρέπει:

- εμφάνιση όλων των θεραπειών του authoritative configuration,
- επιλογή πολλών θεραπειών για τον ίδιο ασθενή,
- χειροκίνητη συμπλήρωση ημέρας, ώρας και provider όπου απαιτείται,
- δυνατότητα ακύρωσης μιας ήδη επιλεγμένης τιμής ώστε το πεδίο να ξαναγίνεται ελεύθερο για automatic completion,
- μερική χειροκίνητη συμπλήρωση και automatic completion μόνο των κενών πεδίων,
- πλήρως automatic proposal όταν ο χρήστης αφήνει όλα τα scheduling πεδία κενά,
- διατήρηση των χειροκίνητων τιμών ως hard constraints,
- ισχυρή soft preference ώστε το Ανακλινόμενο να τοποθετείται ακριβώς πριν ή μετά τη ΦΘ όταν είναι εφικτό,
- μία τελική atomic αποθήκευση όλου του patient program.

### Τρέχον UX micro-fix

Υλοποιήθηκε keyboard clear για τα combo fields της ενεργής θεραπείας:

- `Backspace` ή `Delete` στις **Ημέρες** καθαρίζει το day pattern,
- `Backspace` ή `Delete` στην **Ώρα** καθαρίζει την ώρα,
- `Backspace` ή `Delete` στον **Θεραπευτή / provider** καθαρίζει τον provider,
- η αντίστοιχη τιμή αφαιρείται από το draft και από τη γραμμή σύνοψης,
- το πεδίο θεωρείται ξανά κενό και μπορεί να συμπληρωθεί από την επόμενη Αυτόματη συμπλήρωση.

Το παραπάνω είναι **implemented, pending native Excel verification**.

---

## Τι έχει ολοκληρωθεί μέχρι τώρα

### Application / workbook foundation

- MASTER_SCHEDULE ως operational home και read-only hub.
- Application shell με κρυμμένο Excel chrome στο app mode.
- MASTER toolbar και βασική πλοήγηση.
- DAILY_INPUT authoritative flow.
- THERAPIST_DAILY recurring projection και writeback υποδομή.
- Patient, therapist, student και outpatient registration vertical slices.
- Outpatient scheduling και συμβατότητα με unified schedule.
- MASTER projection, room sorting, infectious styling και outpatient markers.
- Robotic και outpatient recurring visual rules.

### Patient scheduling architecture

- Patient-centric draft model για πολλές θεραπείες.
- Όλες οι θεραπείες προέρχονται από το workbook/configuration και όχι από αυθαίρετο UI subset.
- PATIENT_PLANNER blocks έχουν χαρτογραφηθεί για:
  - ΦΘ
  - Ρομποτικό
  - Ανακλινόμενο
  - Πισίνα
  - Εργο
  - Λογο
  - ΕΦΑ
  - Ψυχ
- Provider pools υπάρχουν για φυσιοθεραπευτές, εργοθεραπευτές, λογοθεραπευτές και ψυχολόγους.
- Workload display διορθώθηκε ώστε recurring load να εκφράζει peak daily load και όχι άθροισμα όλων των ημερών.
- Ο automatic solver γενικεύτηκε πέρα από ΦΘ/Ρομποτικό.
- Κενές ημέρες μπορούν να συμπληρώνονται από το automatic μέσω των day patterns των SETTINGS.
- Manual time/provider/day values αντιμετωπίζονται ως locked constraints και δεν αντικαθίστανται σιωπηρά.
- Providerless θεραπείες μπορούν να προγραμματίζονται χωρίς υποχρεωτικό therapist/provider.
- Recliner/ΦΘ adjacency preference έχει ενσωματωθεί στη βαθμολόγηση του solver ως soft preference.
- Η φόρμα κρατά μία ενεργή θεραπεία για editing και εμφανίζει στη γραμμή της κάθε θεραπείας τις τρέχουσες ημέρες, ώρα και provider.
- Προστέθηκε δυνατότητα keyboard clear στα τρία combo fields ώστε μια manual ή automatic πρόταση να μπορεί να μηδενιστεί και να ξαναδοθεί στον solver.

### Atomic final save

Η τελική αποθήκευση του patient program έχει σχεδιαστεί/υλοποιηθεί με transaction-like συμπεριφορά:

1. validation όλων των επιλεγμένων θεραπειών,
2. resolve όλων των PATIENT_PLANNER headers πριν από οποιοδήποτε write,
3. snapshot όλων των cells που πρόκειται να αλλάξουν,
4. `Application.EnableEvents = False` πριν από το πρώτο write,
5. εγγραφή όλων των assignments,
6. επαλήθευση των τιμών,
7. ένα μόνο `ThisWorkbook.Save`,
8. rollback όλων των touched cells σε περίπτωση σφάλματος,
9. αποκατάσταση του προηγούμενου event state σε κάθε path.

Αυτό αντικαθιστά το παλιό per-treatment save flow που μπορούσε να οδηγήσει σε partial writes και workbook reopen/re-entry.

### Regression coverage που προστέθηκε στο τελευταίο milestone

- generalized automatic scheduling,
- automatic day-pattern completion,
- specialty/provider handling,
- providerless treatment handling,
- recliner adjacency preference,
- patient-program bridge behavior,
- atomic-save safety ordering και rollback contract,
- keyboard clearing για day/time/provider combo selections,
- wiring του combo-clear patch στο scheduling preview installer.

---

## Τρέχον verification gate

Πριν θεωρηθεί ολοκληρωμένο το patient-program milestone πρέπει να περάσουν:

1. focused pytest suite για automatic solver, bridge, form VBA, recurring scheduling και registration slice,
2. tests για keyboard clear των patient-program combos,
3. build νέου preview workbook από το clean V48 base,
4. native Excel smoke test με πολλές ταυτόχρονες θεραπείες,
5. επιβεβαίωση ότι `Backspace`/`Delete` καθαρίζει ημέρες, ώρα και provider και αφαιρεί την τιμή από τη γραμμή σύνοψης,
6. επιβεβαίωση ότι πεδίο που καθαρίστηκε μπορεί να ξανασυμπληρωθεί από την Αυτόματη συμπλήρωση,
7. επιβεβαίωση ότι manual values παραμένουν αμετάβλητες μετά το automatic completion,
8. επιβεβαίωση ότι η τελική αποθήκευση γράφει όλα τα selected treatments μαζί,
9. επιβεβαίωση ότι time cells εμφανίζονται ως `hh:mm`, όχι ως raw Excel serial,
10. επιβεβαίωση ότι το workbook δεν κλείνει/ξανανοίγει κατά την αποθήκευση.

---

## Επόμενα βήματα

### Άμεσα

1. Τρέξιμο focused tests στο branch `feature/outpatient-scheduling`, μαζί με τα νέα combo-clear tests.
2. Build του επόμενου patient-scheduling preview.
3. Native Excel smoke test ειδικά για clear → automatic refill.
4. Διόρθωση τυχόν regressions που θα εμφανιστούν.
5. Native Excel smoke test του πλήρους manual / mixed / automatic flow και της atomic αποθήκευσης.
6. Full regression suite όταν περάσει το focused smoke test.

### Μετά το patient-program milestone

- σύνδεση του τελικού save με refresh όλων των dependent views όπου χρειάζεται,
- τελικός έλεγχος UX της φόρμας και των labels,
- consolidation/cleanup παλιού single-treatment scheduling code που δεν χρησιμοποιείται πλέον,
- ενημέρωση `UI_APPLICATION_ROADMAP.md` και `Documentation/PROJECT_STATUS.md` με το verified status,
- μετάβαση στο επόμενο ανοιχτό roadmap milestone.

---

## Κανόνας ενημέρωσης αυτού του χάρτη

Σε κάθε ουσιαστική αλλαγή θα ενημερώνονται τουλάχιστον:

- **Τι φτιάχνουμε τώρα**
- **Τι έχει ολοκληρωθεί μέχρι τώρα**
- **Επόμενα βήματα**
- **Verification gate**, όταν αλλάζει το test/smoke status

Δεν χαρακτηρίζεται κάτι ως ολοκληρωμένο μόνο επειδή γράφτηκε ο κώδικας. Για milestone που επηρεάζει το workbook/UI, ολοκλήρωση σημαίνει και επιτυχές test/smoke verification.
