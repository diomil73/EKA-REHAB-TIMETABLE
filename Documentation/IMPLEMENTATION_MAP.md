# Implementation Map

## Σκοπός

Το αρχείο αυτό είναι ο ζωντανός χάρτης υλοποίησης του project. Συμπληρώνει το `UI_APPLICATION_ROADMAP.md` και ενημερώνεται σε κάθε ουσιαστικό milestone με τέσσερα πράγματα: **τι φτιάχνουμε τώρα**, **τι έχει ολοκληρωθεί**, **ποιο είναι το verification gate**, **ποια είναι τα επόμενα βήματα**.

Τελευταία ενημέρωση: 2026-10-11

---

## Τι φτιάχνουμε τώρα

### Patient-centric multi-treatment scheduling

Στόχος είναι μία ενιαία ροή προγραμματισμού ασθενούς, χωρίς ξεχωριστό Save ανά θεραπεία.

Η ίδια patient-program φόρμα χρησιμοποιείται τόσο:

- μετά την πρώτη καταχώρηση νέου ασθενούς,
- όσο και μετά από **Βασικό μενού → Επεξεργασία ασθενή → Αποθήκευση αλλαγών**.

Στην edit περίπτωση η φόρμα πρέπει να φορτώνει το ήδη υπάρχον πρόγραμμα του ασθενούς από `PATIENT_PLANNER` και να επιτρέπει αλλαγές χωρίς δεύτερο παράλληλο scheduling UI.

### Provider / όνομα προσωπικού

Η τρέχουσα επιχειρησιακή απόφαση είναι:

- **ΦΘ:** ο θεραπευτής είναι υποχρεωτικός και συμμετέχει στον automatic solver.
- **Ρομποτικό, Εργο, Λογο, Ψυχ:** το scheduling απαιτεί μόνο **ημέρες + ώρα**. Το όνομα είναι **προαιρετικό**.
- Για Ρομποτικό/Εργο/Λογο/Ψυχ ο χρήστης μπορεί να γράψει χειροκίνητα όνομα στη φόρμα. Αν υπάρχει configured λίστα, μπορεί να χρησιμοποιηθεί ως βοήθημα, αλλά το όνομα δεν είναι solver constraint.
- **Ανακλινόμενο, Πισίνα, ΕΦΑ:** day/time scheduling χωρίς provider/name field.
- Η τελική αποθήκευση γράφει τα προαιρετικά ονόματα στις υπάρχουσες στήλες του `PATIENT_PLANNER` όπου αυτές υπάρχουν.

> Μελλοντική optional επέκταση: provider-aware automatic scheduling για Ρομποτικό/Εργο/Λογο/Ψυχ μπορεί να προστεθεί αργότερα, χωρίς να γίνει ποτέ υποχρεωτικό για την εύρεση timeslot.

### Duration-safe scheduling

Ο solver πλέον πρέπει να θεωρεί πραγματικά χρονικά διαστήματα και όχι μόνο ίδια start times:

- standard treatments, συμπεριλαμβανομένης της ΦΘ: **45 λεπτά**,
- Ανακλινόμενο: **30 λεπτά**,
- δύο θεραπείες του ίδιου ασθενούς δεν επιτρέπεται να επικαλύπτονται σε κοινές ημέρες,
- ακριβώς συνεχόμενα intervals επιτρέπονται.

Παράδειγμα: ΦΘ `09:15-10:00` αποκλείει Ανακλινόμενο `09:00-09:30` και `09:30-10:00`. Αν οι διαθέσιμες recliner ώρες είναι `08:30, 09:00, 09:30, 10:00`, η καλύτερη adjacent επιλογή είναι `10:00`. Αν χρειαστεί να μπει πριν, η τελευταία συμβατή διαθέσιμη είναι `08:30` επειδή `08:45` δεν υπάρχει ως slot.

### MASTER treatment rendering

Το `MASTER_SCHEDULE` δεν πρέπει να εμφανίζει απλώς raw Excel time serial ούτε να εφαρμόζει γενικό `NumberFormat = "hh:mm"` στις treatment columns.

Κάθε treatment cell είναι **derived presentation** από τα authoritative πεδία του `PATIENT_PLANNER` και πρέπει να εμφανίζει:

```text
09:15
Δε-Τε-Πα
Όνομα θεραπευτή
```

Η τρίτη γραμμή εμφανίζεται μόνο όταν υπάρχει όνομα. Για providerless θεραπείες εμφανίζονται ώρα + ημέρες. Η προβολή πρέπει να ανανεώνεται μετά από patient-program save και μετά από patient registry edit/rebuild.

---

## Τι έχει ολοκληρωθεί μέχρι τώρα

### Application / workbook foundation

- `MASTER_SCHEDULE` ως operational home και read-only hub.
- Application shell με κρυμμένο Excel chrome στο app mode.
- MASTER toolbar και βασική πλοήγηση.
- `DAILY_INPUT` authoritative flow.
- `THERAPIST_DAILY` recurring projection και writeback υποδομή.
- Patient, therapist, student και outpatient registration vertical slices.
- Outpatient scheduling και unified base schedule compatibility.
- MASTER projection, room sorting, infectious styling και outpatient markers.
- Robotic και outpatient recurring visual rules.

### Patient scheduling architecture

- Patient-centric draft model για πολλές θεραπείες.
- Όλες οι θεραπείες προέρχονται από authoritative workbook configuration.
- `PATIENT_PLANNER` blocks για ΦΘ, Ρομποτικό, Ανακλινόμενο, Πισίνα, Εργο, Λογο, ΕΦΑ και Ψυχ.
- Recurring workload ως peak daily load, όχι άθροισμα recurrence.
- Generalized automatic solver πέρα από ΦΘ/Ρομποτικό.
- Automatic day-pattern completion όταν οι ημέρες μένουν κενές.
- Manual day/time values ως fixed constraints.
- Provider constraint μόνο στη ΦΘ.
- Ρομποτικό/Εργο/Λογο/Ψυχ προγραμματίζονται χωρίς υποχρεωτικό provider.
- Providerless treatment support.
- Μία ενεργή θεραπεία για editing, με persistent draft ανά θεραπεία.
- `Backspace`/`Delete` clear για scheduling πεδία.
- Diagnostic failure reasons για πραγματικές scheduling συγκρούσεις.
- Atomic final save με snapshot, events-off-before-write, verification, single workbook save και rollback.

### Νεότερο batch που έχει υλοποιηθεί σε κώδικα και περιμένει native verification

- Separation `supports_provider` από `requires_provider`.
- Optional χειροκίνητο όνομα για Ρομποτικό/Εργο/Λογο/Ψυχ.
- Optional name δεν περνά ως constraint στον automatic solver.
- Optional name αποθηκεύεται στη σχετική provider/name στήλη του `PATIENT_PLANNER`.
- Editable provider combo για optional names, ακόμη και όταν η configured λίστα είναι κενή.
- `Delete`/`Backspace` καθαρίζει και typed free-text provider value.
- `config-vba` φορτώνει υπάρχον πρόγραμμα ως `CURRENT` rows.
- Η patient-program φόρμα προφορτώνει υπάρχουσες ημέρες/ώρες/ονόματα.
- Η **Επεξεργασία ασθενή** μετά την αποθήκευση registry αλλαγών καλεί `OpenPatientScheduling patientId` και οδηγεί στην ίδια scheduling φόρμα.
- Duration-aware patient conflict logic: standard 45', Ανακλινόμενο 30'.
- Recliner/ΦΘ adjacency scoring με πραγματικό gap μεταξύ intervals.
- Νέο `PatientProgramMasterRender` module χτίζει τα treatment cells του MASTER από ώρα + ημέρες + όνομα, αντί για global time number format.
- Το patient-program save καλεί MASTER refresh πριν από το τελικό workbook save.
- Το patient edit flow επαναφέρει το composite treatment rendering μετά από registry/master rebuild.

### V65 installer hardening και MASTER regression

Το πρώτο V65 build σταμάτησε με Excel/VBE COM error `-2146828213`, που αντιστοιχεί σε **Run-time error 75 / Path/File access error**, μέσα στο scheduling UI installation stage.

Έγιναν δύο installer διορθώσεις:

- το `frmEditPatient` δεν αφαιρείται και δεν ξαναδημιουργείται πλέον όταν υπάρχει ήδη. Ο installer έγινε idempotent,
- ο scheduling UI installer απέκτησε stage-specific diagnostics.

Στο επόμενο native smoke test εντοπίστηκε δεύτερο regression: αντί για πραγματική ώρα εμφανίστηκε κυριολεκτικά `hh:mm` και δεν εμφανίστηκαν ημέρες/θεραπευτής. Η αιτία ήταν η λανθασμένη υπόθεση ότι το MASTER χρειάζεται μόνο `NumberFormat = "hh:mm"` πάνω στις treatment columns. Αυτό αφαιρέθηκε και αντικαταστάθηκε από explicit composite rendering από `PATIENT_PLANNER`.

### Τελευταίο verified gate

Το προηγούμενο focused gate έχει επιβεβαιωθεί από τον χρήστη:

**46 passed in 2.54s**

Αυτό αφορούσε την προηγούμενη provider-optional αρχιτεκτονική πριν από το νεότερο batch edit/preload/manual-name/duration/MASTER-render changes.

---

## Τρέχον verification gate

Πριν θεωρηθεί verified το νεότερο batch χρειάζονται:

1. focused pytest για draft, solver, bridge, program form, edit flow, combo clear, MASTER renderer και scheduling slice,
2. νέο preview build από clean V48 base,
3. έλεγχος ότι κάθε treatment cell στο MASTER εμφανίζει πραγματική ώρα και ημέρες, όχι literal `hh:mm`,
4. έλεγχος ότι ΦΘ εμφανίζει και θεραπευτή και ότι Ρομποτικό/Εργο/Λογο/Ψυχ εμφανίζουν προαιρετικό όνομα όταν έχει συμπληρωθεί,
5. έλεγχος ότι providerless θεραπείες εμφανίζουν ώρα + ημέρες,
6. έλεγχος ΦΘ `09:15` + Ανακλινόμενο ώστε να μην επιτρέπεται overlap και να προτιμάται `10:00` έναντι `08:30` όταν είναι διαθέσιμο,
7. έλεγχος ότι optional όνομα μπορεί να γραφτεί χειροκίνητα χωρίς να μπλοκάρει automatic completion,
8. έλεγχος ότι το optional όνομα αποθηκεύεται και ξαναφορτώνεται,
9. διαδρομή **Βασικό μενού → Επεξεργασία ασθενή → αλλαγή θαλάμου/στοιχείων → Αποθήκευση → ίδια patient-program φόρμα**,
10. επιβεβαίωση ότι το MASTER composite rendering παραμένει σωστό και μετά από αλλαγή θαλάμου/rebuild,
11. έλεγχος atomic save μετά από edit υπάρχοντος προγράμματος.

---

## Επόμενα βήματα

### Άμεσα

1. Τρέξιμο focused regression suite για MASTER renderer + νέο patient-program batch.
2. Build επόμενου preview από `MASTER_TOOLBAR_SMOKE_V48.xlsm`.
3. Native smoke test πρώτα μόνο σε μία ΦΘ και μία providerless θεραπεία για να επιβεβαιωθεί το cell rendering.
4. Μετά smoke test duration-safe adjacency, edit existing program και optional manual names.
5. Full regression suite όταν περάσει το native smoke test.

### Μετά το patient-program milestone

- refresh dependent views μετά το final save όπου χρειάζεται,
- τελικός UX έλεγχος labels και edit affordances,
- cleanup παλιού single-treatment scheduling code,
- ενημέρωση `UI_APPLICATION_ROADMAP.md` και `Documentation/PROJECT_STATUS.md` με verified status,
- μετάβαση στο επόμενο ανοιχτό roadmap milestone.

---

## Κανόνας ενημέρωσης αυτού του χάρτη

Σε κάθε ουσιαστική αλλαγή ενημερώνονται τουλάχιστον:

- **Τι φτιάχνουμε τώρα**
- **Τι έχει ολοκληρωθεί μέχρι τώρα**
- **Τρέχον verification gate**
- **Επόμενα βήματα**

Κανένα milestone δεν χαρακτηρίζεται ολοκληρωμένο μόνο επειδή γράφτηκε κώδικας. Χρειάζεται regression verification και, όταν αφορά Excel/VBA συμπεριφορά, native workbook smoke test.