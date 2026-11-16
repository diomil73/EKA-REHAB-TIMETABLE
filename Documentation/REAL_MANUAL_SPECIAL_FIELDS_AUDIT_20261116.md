# Πραγματικό χειροκίνητο A:V: ανεξάρτητος έλεγχος F/U/V (2026-11-16)

**Status:** Παραδόθηκε read-only κώδικας ελέγχου, όχι πραγματικό Excel DEMO ούτε έγκριση αυτόματης εισαγωγής. Το εργαλείο λειτουργεί τοπικά σε *φρέσκο πραγματικό, ιδιωτικό* `.xlsx` χειροκίνητο πρόγραμμα, χωρίς να απαιτεί ακόμη εγκεκριμένο `review.xlsx` ή κενό `.xlsm`.

**Εκτέλεση μόνο από εξουσιοδοτημένο Windows/Python περιβάλλον:**

```powershell
python .\Python\tools\audit_real_manual_specials.py `
  --programme "C:\PRIVATE_APPROVED_DIOMIL\manual_current.xlsx"

$specialAuditExit = $LASTEXITCODE
if ($specialAuditExit -eq 2) { throw "SAFETY STOP: source schema or read-only inspection invalid" }
if ($specialAuditExit -eq 3) {
    "REVIEW REQUIRED: genuine F/U/V treatment cells are still blocked"
}
if ($specialAuditExit -eq 0) {
    "F/U/V are empty; full original-source + clinical review still required"
}
```

Η διαδρομή είναι υποθετικό παράδειγμα. **Μην μεταφέρετε** αρχεία ή screenshots ασθενών στο δημόσιο GitHub. Χρησιμοποιήστε εξουσιοδοτημένο, ιδιωτικό αποθηκευτικό χώρο. Η εντολή **δεν γράφει, δεν αντιγράφει, δεν μετατρέπει και δεν δημιουργεί XLSM**. Ελέγχει SHA-256 πριν/μετά. Σε σφάλμα δεν εμφανίζει όνομα ασθενούς, ιατρού, θάλαμο, ID, ακατέργαστη τιμή κελιού ή πλήρη ιδιωτική διαδρομή.

## Ακριβής επιθεώρηση

| Κελί στο αρχικό A:V | Επιβεβαιωμένο | Τι αναφέρει | Τι ΔΕΝ κάνει |
| --- | --- | --- | --- |
| F | Ανακλινόμενο, ΦΘ-σχετικό, διάρκεια 30 λεπτά, στο ίδιο κελί μόνο ώρα | Εγγραφές, μηχανικά αναγνώσιμες ώρες, ασαφείς τιμές/τύποι Excel | **Δεν εφευρίσκει ημέρες**, διάρκεια σε γενικό slot, therapist ή ειδική χωρητικότητα |
| U | Παλαιό απογευματινό, μη επιλέξιμο για auto-import | Πλήθος συμπληρωμένων πεδίων | Δεν τα αντιγράφει σε πρωινή ΦΘ ή απογευματινό πρόγραμμα |
| V | Πισίνα, πρωινή θεραπεία, ίδια υπάρχουσα ειδικότητα `Πισίνα` | Πλήθος, αναγνώσιμες ώρες, ασαφείς τιμές/τύποι | **Δεν εφευρίσκει ημέρες**, provider, χωρητικότητα ή δεύτερο treatment token |

Οι τύποι Excel στα F/U/V **μετρώνται ως εκκρεμείς**, ακόμα και αν το `data_only=True` δεν διαθέτει αποθηκευμένη τιμή. Καταγράφονται επίσης ειδικές εγγραφές σε γραμμή χωρίς ασθενή, αντί να αγνοηθούν. Δεν γίνεται σιωπηλή μερική εισαγωγή. Όλα τα αναγνώσιμα F/V εξακολουθούν να απαιτούν εγκεκριμένη ημέρα/κανόνες και παραμένουν blockers.

**Return code:** `0` σημαίνει μόνο ότι δεν υπάρχουν συμπληρωμένα F/U/V, `3` σημαίνει pending ειδικές εγγραφές (αναμενόμενο σε πραγματικό πρόγραμμα), `2` σημαίνει αναγνώσιμο/δομικό safety stop. **Κανένα από αυτά δεν σημαίνει πλήρη έγκριση όλων των θεραπειών/ασθενών ή native Excel pass.**

## Τι απομένει για άρση blockers

1. Τοπική επιθεώρηση **του ίδιου πραγματικού χειροκίνητου snapshot** και εξουσιοδοτημένη καταχώριση ημέρας/ώρας/παρόχου ανά F ή V, χωρίς συναγωγές μόνο από το ότι εμφανίζεται ρολόι. Για το F: επικύρωση του 30λεπτου interval και κανόνων γειτνίασης/αναπλήρωσης, πριν χρησιμοποιηθεί ο γενικός daily/capacity engine.
2. Το U **δεν εισάγεται αυτόματα**, σύμφωνα με την ρητή οδηγία. Το κανονικό απογευματινό σύστημα είναι ξεχωριστό, αρχικά κενό, με εξουσιοδοτημένη εισαγωγή νέων ωρών/ημερών/φυσικοθεραπευτών. Η ύπαρξη U πρέπει να παρουσιαστεί στη σύγκριση ως χωριστή εκκρεμότητα, όχι ως κρυμμένο/χαμένο treatment.
3. Εγκεκριμένο `review`, κλινικά fields, εγκεκριμένο καθαρό XLSM. Η υπάρχουσα real-data CLI παραμένει **fail-closed** όσο εκκρεμούν αυτές οι πύλες.
4. Ιδιωτική native COM κατασκευή + save/close/reopen + πλήρης ΔΕ–ΠΑ 1:1 σύγκριση με το χειροκίνητο και λειτουργικές δοκιμές χρήστη.

## Κώδικας και τεκμηρίωση

- `Python/rehab_excel/trial_special_fields_audit.py`: read-only scanner, πλήρη μη ευαίσθητα counts, time parsing μόνο για διάγνωση, διπλή raw/cached ανάγνωση τύπων, αμετάβλητο πρωτότυπο.
- `Python/tools/audit_real_manual_specials.py`: ασφαλής συγκεντρωτική CLI χωρίς διαρροή προσωπικών δεδομένων.
- `Tests/test_trial_special_fields_audit.py`: synthetic tests για ακριβείς μετρήσεις, formula-only cells, orphan rows, αλλοιωμένα headers, απαγόρευση μεταβολών και ασφαλή έξοδο.

**Δεν έχει γίνει** σάρωση πραγματικού ιδιωτικού αρχείου ή επιβεβαίωση αποδεκτού πραγματικού DEMO. Όταν εκτελεστεί σε ιδιωτικό, πραγματικό αρχείο, καταγράφονται στις ζωντανές αναφορές μόνο οι συνολικοί αριθμοί F/U/V και ο κωδικός επιστροφής, χωρίς ονόματα ή κελιά. Το πλήρες authoritative handoff παραμένει `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md`.


## 2026-11-16 — Mandatory F/U/V source gate before real DEMO seed

`Python/tools/prepare_real_data_practice_demo.py` now calls `inspect_real_manual_special_fields(programme)` **before reading approved clinical values, building source plans or invoking native Excel**. Any populated F reclining, U legacy afternoon or V morning pool field (including Excel formula-only content with empty cache), or an orphaned source row, gives a privacy-safe **SAFETY STOP**. This ensures that even a stale review workbook cannot silently hide real original-source special therapies. F/V clock parseability is diagnostic only: **days/duration-aware scheduling and authoritative source treatment review are still required**. U remains expressly non-importable from the historical column; separately entered authorized afternoon appointments belong to their own programme. The standalone command `python Python\tools\audit_real_manual_specials.py --programme "<PRIVATE_CURRENT_MANUAL>.xlsx"` reports only aggregate counts and returns 3 for blocked real specials, 0 if no F/U/V cells are populated, 2 on structural safety failure. No original workbook or native XLSM was changed in the GitHub development environment.

**Tests:** `Tests/test_prepare_real_data_practice_demo.py` adds explicit refusal of preflight and native build before clinical approval/Excel on pending F/U/V/orphan formula cells, alongside `Tests/test_trial_special_fields_audit.py`. Last fully green complete regression at earlier source commit `95157af7dbb32c6cb8aca210c503ffbb7f63be17` was 1,991 passing Windows/Linux and 4/4 CI jobs GREEN (run `43804737124`). This integrated new hard-stop code requires its own fresh CI confirmation. Real private source inspection, source review, native Excel candidate and user acceptance remain **PENDING**. The new hard-stop is deliberate and must NOT be bypassed with invented patients or source omissions.
