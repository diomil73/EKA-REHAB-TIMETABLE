# Real outpatient current-source check using established Diomil sheets (2026-11-16)

**Status:** Source-only diagnostic delivered. **It does not import any outpatient into a real XLSM and does not certify the business demo.** The existing operational UI, source columns and treatment-cell layout are preserved.

## Required local input

A separately authorized PRIVATE current outpatient export **outside Git**: an `.xlsx` workbook containing the already established Diomil sheet names:

- `PATIENTS`, original columns 1–6 in exactly this order: `PatientID`, `Θάλαμος`, `Ασθενής`, `Λοιμώδης`, `Κατάσταση`, `PatientType`. Existing additional metadata columns may remain. All outpatient rows must be explicitly typed `Εξωτερικός` (or other values already accepted by the production registry reader); do not infer from names or rooms. A stable PatientID and current identity must be independently approved by authorized staff.
- `OUTPATIENT_SCHEDULE`, original six columns: `PatientID`, `Ασθενής`, `Θεραπεία`, `Ώρα`, `Ημέρες`, `Θεραπευτής`. These are not new columns or a parallel appointment schema.

**A real current manual export must actually be available**. This diagnostic cannot conjure it from the inpatient A:V manual workbook. No source has been connected from the user's local computer to GitHub.

## Private read-only command

```powershell
python .\Python\tools\audit_real_outpatient_source.py `
  --source "C:\PRIVATE_APPROVED_DIOMIL\outpatients_current.xlsx"

if ($LASTEXITCODE -ne 0) {
    throw "SAFETY STOP: outpatient source cannot be admitted"
}
```

Private path is an example. **Do not upload patient information or paste individual lines into the chat/GitHub.** The CLI prints only four aggregate numbers (patients, recurring treatment rows, Monday–Friday placements and distinct treatments), a source-unchanged proof and explicit **import PENDING** status. Errors expose only their generic class, never a private identifier/path.

The implementation `Python/rehab_excel/trial_private_outpatient_audit.py` reuses the EXISTING production `read_patient_registry`, `read_outpatient_schedule` and `parse_day_pattern`, rather than creating another representation. It verifies source SHA-256 before/after, prohibits Excel formula-only input in the six basic source fields, missing/unknown/duplicate patient IDs, invalid patient type, unsupported robotic outpatient therapies, weekends, a missing physiotherapist for ΦΘ, duplicate weekday/time/treatment slots across day patterns, and active outpatient rows with no corresponding recurring treatment. This source-only auditor is intentionally stricter than the legacy reader in incomplete/ambiguous cases because real-data migration must not silently lose an appointment.

**Return values:** 0 = aggregate/read-only outpatient audit passes, 2 = safety stop. Zero **does not** mean native integration or admission is accepted. `Tests/test_trial_private_outpatient_audit.py` covers the paths using only fictional patients and providers.

## Next actual engineering task

Join the independently approved, SHA-bound current outpatient source with the separate current inpatient A:V source in the **existing** native DEMO seed transaction. Validate patient ID uniqueness across both, add outpatient rows to the existing `PATIENTS` and `OUTPATIENT_SCHEDULE` sheets without guessing source values, and independently reopen through the existing unified production schedule reader. Check full exact inpatients + outpatients per-treatment/day/time/provider and live absence/replacement/restore. F/U/V special source holds must still be resolved independently, and the manual source remains the authoritative operational record until signoff. None of this is proven by the source-only audit.

See `Documentation/REAL_DATA_DEMO_ACCEPTANCE_AND_FULL_HANDOFF_20261116.md` and `Documentation/REAL_DATA_DEMO_OPERATOR_CLI_20261116.md`.
