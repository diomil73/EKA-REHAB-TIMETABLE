# Operator UI Rules

Last updated: 2026-10-25

## Purpose

The application must present operational actions to normal users as an application, not as a Python/VBA/Excel development environment.

## Locked rules

1. Normal user workflows must not display console/terminal/Command Prompt windows.
2. Background Python or worker processes must run hidden.
3. User-facing messages must not expose implementation terms such as backend, worker, preview, hash, COM, VBA module names, file-lock mechanics, or internal technical tables.
4. Success messages should be short and state only the completed operational result.
5. Error messages should tell the operator what to do next. They should not explain internal processing unless an administrator/developer diagnostic tool is being used explicitly.
6. Confirmation dialogs should describe the business consequence of the action, not the internal execution chain.
7. Technical diagnostics remain available in developer tools/logs/tests, not in normal operator UI.
8. These rules apply to all future user-facing workflows, not only `Κλείσιμο ημέρας`.

## Close-day example

Preferred normal-user language:

- Confirmation: `Να κλείσει οριστικά η επιλεγμένη ημέρα; Μετά το κλείσιμο η ημέρα θα περάσει στο ιστορικό.`
- Success: `Η ημέρα έκλεισε επιτυχώς.`
- Recoverable failure: `Η ημέρα δεν έκλεισε. Ελέγξτε ότι δεν υπάρχουν εκκρεμείς συνεδρίες και δοκιμάστε ξανά.`
- Technical/repeated failure: `Αν το πρόβλημα συνεχιστεί, επικοινωνήστε με τον διαχειριστή.`

Do not expose the verified-preview, authoritative-commit, source-hash, worker, or reopen mechanics to the operator.
