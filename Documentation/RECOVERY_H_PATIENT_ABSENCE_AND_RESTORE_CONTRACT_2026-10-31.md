# Recovery H patient-absence + restore contract — 2026-10-31

## Native H evidence

Artifact under test: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_H.xlsm`.

Operator-confirmed patient-absence behavior:

- recurring treatment remains visible in `MASTER_SCHEDULE`,
- affected treatment cell is red and struck through,
- `Κατάσταση` no longer remains `παρών`,
- the absence reason is written in the status column,
- defect: the status/reason cell itself was also struck through, reducing readability.

H is frozen as native evidence and must not be repaired in place. Close-day native testing remains blocked.

## Bounded presentation fix

The installed DAILY_INPUT action already had separate source tags for the struck treatment cell and the patient-status reason. The patient-status tag set red font but did not explicitly clear inherited strike-through formatting.

Recovery decision:

- `master_daily_non_active_overlay`: red + strike-through,
- `master_daily_patient_status_overlay`: red + explicit `Strikethrough = False`,
- semantic readback must verify both red color and `Strikethrough = False`.

This is presentation-only and does not change daily-state semantics.

## Recovered daily-change / restore UX contract

Historical/current repository evidence confirms the intended operator model below `THERAPIST_DAILY`:

- a visible section `Αλλαγές για την ...`,
- each change stays as a historical row,
- columns include patient, `Από`, `Προς`, reason, status and action,
- human-readable transition such as `Ασθενής Χ: Θεραπευτής Ψ 09:00 -> Θεραπευτής Α 10:30`,
- one-row `Επαναφορά`,
- checkbox selection for active rows,
- `Επαναφορά επιλεγμένων`,
- `Επαναφορά όλων`,
- restored rows remain visible with status `Επαναφέρθηκε`; active rows show `Ενεργή αλλαγή`.

This UX contract is retained.

## Architecture decision

Do **not** revive the obsolete close/reopen restore path.

The later in-session recovery wrapper removed terminal/close-reopen behavior, but still creates a separate restore preview and synchronizes an operational sheet set back into the open workbook. That remains older than the current production mutation contract.

For the next recovery slice, preserve the ledger UX while replacing the restore backend with the current architecture:

1. validate requested replacement IDs / active set,
2. prepare a read-only restore/recompute payload,
3. mutate the already-open workbook in place,
4. update `DAILY_REPLACEMENT_FACTS` and `DAILY_CHANGE_LEDGER` atomically,
5. refresh `THERAPIST_DAILY`, `MASTER_SCHEDULE` and daily statistics from effective truth,
6. perform semantic readback,
7. rollback on any mismatch,
8. save the same workbook only after successful verification.

No production workbook mutation and no repair of H in place.

## Next gate

Before building a fresh `I` candidate:

- focused regression for non-struck MASTER status reason,
- restore/ledger contract tests remain green,
- design/implementation of the in-place restore bridge must reject close/reopen and preview-sync ownership,
- only then install the ledger/restore action into the canonical builder and promote the application contract if a new component is required.
