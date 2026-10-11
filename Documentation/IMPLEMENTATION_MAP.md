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

Το `MASTER_SCHEDULE` δεν πρέπει να εμφανίζει raw Excel time serial ούτε να εφαρμόζει γενικό `NumberFormat = "hh:mm"` στις treatment columns.

Κάθε treatment cell είναι derived presentation από τα authoritative πεδία του `PATIENT_PLANNER` και πρέπει να εμφανίζει:

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

Το πρώτο V65 build σταμάτησε με Excel/VBE COM error `-2146828213` μέσα στο scheduling UI installation stage.

Έγιναν δύο installer διορθώσεις:

- το `frmEditPatient` δεν αφαιρείται και δεν ξαναδημιουργείται πλέον όταν υπάρχει ήδη. Ο installer έγινε idempotent,
- ο scheduling UI installer απέκτησε stage-specific diagnostics.

Στο επόμενο native smoke test εντοπίστηκε δεύτερο regression: αντί για πραγματική ώρα εμφανίστηκε κυριολεκτικά `hh:mm` και δεν εμφανίστηκαν ημέρες/θεραπευτής. Η λανθασμένη global-formatting προσέγγιση αφαιρέθηκε και αντικαταστάθηκε από explicit composite rendering από `PATIENT_PLANNER`.

Το focused regression μετά το stale-test fix πέρασε:

**38 passed in 3.25s**

### V66 legacy VBA compile regression

Στο native V66 εμφανίστηκε **VBA Syntax Error** σε κώδικα της παλιάς single-treatment φόρμας. Τα markers `cboTreatment` και `LoadRecurringSchedulingOptions` αποδεικνύουν ότι το workbook εξακολουθούσε να περιέχει legacy scheduling code, παρότι ο νέος patient-program installer είχε εκτελεστεί.

Ο retired μηχανισμός είναι:

- module `PatientSchedulingRuntime`,
- παλιό `frmPatientScheduling` με `cboTreatment`,
- `LoadRecurringSchedulingOptions`,
- `SavePatientRecurringAssignment`.

Ο νέος authoritative μηχανισμός είναι:

- module `PatientProgramRuntime`,
- `frmPatientScheduling` με `lstTreatments`,
- `LoadPatientProgramConfig`,
- `CompletePatientProgram`,
- `SavePatientProgramAtomic`.

Προστέθηκε **VBProject read-back verification πριν από το workbook save**. Το preview build πλέον απορρίπτεται αν:

- υπάρχει ακόμη `PatientSchedulingRuntime`,
- λείπει `PatientProgramRuntime` ή `frmPatientScheduling`,
- λείπουν τα νέα markers `lstTreatments`, `SavePatientProgramAtomic`, `BuildAutomaticRequestArguments`, `LoadPatientProgramConfig`, `CompletePatientProgram`,
- ή βρεθούν legacy markers `cboTreatment`, `LoadRecurringSchedulingOptions`, `SavePatientRecurringAssignment`.

Έτσι δεν επιτρέπεται πλέον να παραχθεί preview που είναι «μισό παλιό, μισό νέο» χωρίς να σταματήσει ο builder.

### V67 MASTER-render installer hardening

Στο πρώτο V67 build ο builder σταμάτησε σωστά πριν το save με:

`MASTER render patch marker missing: patient-program save refresh`

Η αιτία ήταν line-ending mismatch από το Excel VBE: `CRLF` στο `CodeModule.Lines(...)` έναντι `LF` στα Python patch markers. Ο `_patch_code_module` έγινε line-ending agnostic και απέκτησε regression tests για VBE-style `CRLF` και idempotent already-patched input.

Στο επόμενο V67 build το patient-program save hook πέρασε, αλλά ο builder σταμάτησε στο δεύτερο hook:

`MASTER render patch marker missing: patient-edit refresh`

Αυτό έδειξε ότι και το edit hook ήταν υπερβολικά δεμένο με ακριβή γειτονική whitespace/blank-line μορφή γύρω από το multiline `UpdatePatientInWorkbook` call.

Ο patient-edit hook έγινε πλέον **procedure-scoped**:

- εντοπίζει αποκλειστικά το `Private Sub cmdSave_Click() ... End Sub`,
- δεν βασίζεται σε άλλο occurrence του `Trim$(cboStatus.Value)` μέσα στη φόρμα,
- εισάγει `RefreshMasterPatientProgram patientId` αμέσως πριν από το success `MsgBox`, δηλαδή μετά την επιτυχή `UpdatePatientInWorkbook`,
- παραμένει idempotent αν το refresh υπάρχει ήδη,
- διατηρεί explicit safety-stop αν λείπει το `cmdSave_Click` ή το αναμενόμενο success marker.

Προστέθηκαν regression tests που καλύπτουν CRLF, διαφορετικό spacing και duplicate-looking `cboStatus` references εκτός `cmdSave_Click`.

### V67/V68 first-save runtime blocker

Μετά το installer hardening το preview άνοιξε και η νέα patient-program φόρμα λειτούργησε μέχρι το **πρώτο Save**. Αρχικά εμφανίστηκε γενικό `Subscript out of range`.

Τα runtime stage diagnostics απομόνωσαν το failure σε:

`MASTER refresh failed at locating MASTER row: Subscript out of range`

Η αιτία βρέθηκε στη helper `FirstLine()`. Το `FindMasterPatientRow` σαρώνει τις γραμμές του MASTER και περνά το κελί ασθενούς από `FirstLine`. Σε κενές MASTER rows το `Split("", vbLf)` μπορούσε να προκαλέσει `Subscript out of range`.

Η `FirstLine()` πλέον κάνει explicit empty-cell guard πριν από το `Split`.

### V69 first-save verified και automatic edit revalidation

Το native V69 smoke ολοκλήρωσε επιτυχώς την πρώτη αποθήκευση και το MASTER rendering. Ο προηγούμενος first-save blocker θεωρείται κλειστός.

Κατά χειροκίνητη αλλαγή ώρας σε άλλη ειδικότητα και επανεκτέλεση της **Αυτόματης συμπλήρωσης**, εμφανίστηκε μη κρίσιμο false alarm για τη ΦΘ, ενώ η τελική **Αποθήκευση αλλαγών** περνούσε κανονικά. Η αιτία ήταν ότι ο therapist candidate engine έβλεπε την ήδη αποθηκευμένη ΦΘ του ίδιου ασθενούς ως κατειλημμένο session όταν επανεπικύρωνε το draft.

Η διόρθωση είναι edit-aware:

- για τον ίδιο `PatientID` και την ίδια θεραπεία εντοπίζεται το υπάρχον `base_entry_id`,
- αυτό περνά ως `ignore_session_id` στον therapist candidate engine,
- η ήδη υπάρχουσα ΦΘ δεν συγκρούεται πλέον με τον εαυτό της,
- οι πραγματικές συγκρούσεις με άλλους ασθενείς/συνεδρίες παραμένουν ενεργές.

Προστέθηκε regression test που απαιτεί το υπάρχον slot του ίδιου ασθενούς να παραμένει διαθέσιμο, ενώ slot άλλου ασθενούς στον ίδιο θεραπευτή παραμένει δεσμευμένο.

---

## Τρέχον verification gate

Πριν θεωρηθεί ολοκληρωμένο το patient-program milestone χρειάζονται:

1. focused pytest για το νέο edit-aware therapist self-conflict fix,
2. native έλεγχος ότι αλλαγή ώρας σε άλλη ειδικότητα + **Αυτόματη συμπλήρωση** δεν εμφανίζει πλέον ψευδές error για αμετάβλητη ΦΘ,
3. έλεγχος ΦΘ `09:15` + Ανακλινόμενο ώστε να μην επιτρέπεται overlap και να προτιμάται `10:00` έναντι `08:30` όταν είναι διαθέσιμο,
4. έλεγχος ότι optional όνομα για Ρομποτικό/Εργο/Λογο/Ψυχ αποθηκεύεται και ξαναφορτώνεται,
5. διαδρομή **Βασικό μενού → Επεξεργασία ασθενή → αλλαγή θαλάμου/στοιχείων → Αποθήκευση → ίδια patient-program φόρμα**,
6. έλεγχος ότι υπάρχον πρόγραμμα φορτώνεται και μπορεί να τροποποιηθεί και να αποθηκευτεί atomic,
7. επιβεβαίωση ότι το MASTER composite rendering παραμένει σωστό μετά από edit/rebuild,
8. full regression suite.

---

## Επόμενα βήματα

### Άμεσα

1. Τρέξιμο focused tests για `workbook_scheduling_candidates`, recurring scheduling και automatic patient schedule.
2. Native retest του μικρού false-alarm scenario της ΦΘ.
3. Smoke test edit υπάρχοντος προγράμματος και optional manual names.
4. Smoke test duration-safe ΦΘ/Ανακλινόμενο adjacency.
5. Full regression suite όταν περάσουν τα παραπάνω.

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