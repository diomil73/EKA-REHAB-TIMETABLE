# Rehab Center System

Σύστημα προγραμματισμού και καθημερινής λειτουργίας αποκατάστασης/φυσικοθεραπείας.

## Στόχος

Το Excel παραμένει το κύριο περιβάλλον εργασίας των χρηστών. Η Python αναλαμβάνει σταδιακά τη σύνθετη επιχειρησιακή λογική, ώστε οι αλλαγές να ελέγχονται με tests και να μειώνονται οι παρενέργειες σε μορφοποίηση, VBA και φύλλα εργασίας.

## Baseline

- `Rehab_Center_System_v27_1.xlsm`: immutable baseline / έκδοση αναφοράς.
- `Rehab_Center_System_v28_WORKING.xlsm`: working copy για ανάπτυξη.
- Το baseline δεν τροποποιείται.

## Αρχιτεκτονική

- **Excel**: UI, προβολές, φόρμες, μορφοποίηση, κουμπιά, αναφορές.
- **Python**: availability, conflicts, replacements, workload, validation.
- **VBA**: παραμένει όσο χρειάζεται και αποσύρεται σταδιακά, όχι με μία μεγάλη μετατροπή.

## Κανόνας ανάπτυξης

1. Ορίζουμε τον κανόνα.
2. Γράφουμε test.
3. Υλοποιούμε τον κανόνα σε Python.
4. Ελέγχουμε ότι το test περνά.
5. Μόνο μετά συνδέουμε τη λογική με το Excel.

## Πρώτο Python milestone

Το πρώτο milestone καλύπτει:

- βασικά μοντέλα `Patient`, `Therapist`, `Session`,
- ημερήσιες απουσίες ασθενών/θεραπευτών,
- υπολογισμό πραγματικής διαθεσιμότητας θεραπευτή,
- ελευθέρωση slot όταν απουσιάζει ο ασθενής.

## Repository layout

```text
Documentation/
Python/
  rehab_core/
  rehab_excel/
  tools/
Tests/
README.md
CHANGELOG.md
pyproject.toml
```


## Excel write-back

Η Python δεν αποθηκεύει το production `.xlsm` με `openpyxl`. Το πραγματικό write-back γίνεται μόνο μέσω Microsoft Excel σε Windows και μόνο σε νέο preview copy, αφού περάσουν source/layout/render safety checks.
