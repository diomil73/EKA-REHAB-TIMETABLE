# One-shot native recurring AssignmentID migration on synthetic CC

## Strict scope
The previous successful CC read-only probe and synthetic Save/reopen A are frozen.
This new run generates an entirely new `_DEMO_` .xlsm candidate. It must not
be used on the live CC, MASTER, genuine patient data or an existing candidate.

The installer reads the frozen synthetic CC, plans opaque recurring IDs in
memory, checks that all targeted source cells are still correct and unused,
exclusively copies the source, opens **only the fresh copy** in native Excel
with VBA macros disabled, installs the separate planner identity registry and
optional metadata column G on weekly/outpatient sources, Saves once, closes,
then independently verifies every ID, existing schedules, H8 history and VBA.

## Single PowerShell block

Close all Excel. Run the entire block from the repository root. Do not execute
individual lines after an exception.

```powershell
& {
    $ErrorActionPreference = "Stop"

    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: git pull" }

    python -m pytest Tests/test_stable_recurring_ids.py Tests/test_recurring_id_migration_plan.py Tests/test_install_synthetic_recurring_ids_native.py Tests/test_recurring_change_preview.py -q
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: source tests" }

    $source = ".\Excel\previews\UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm"
    $candidate = ".\Excel\previews\UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm"
    if (!(Test-Path -LiteralPath $source -PathType Leaf)) { throw "SAFETY STOP: CC missing" }
    if (Test-Path -LiteralPath $candidate) { throw "SAFETY STOP: candidate already exists" }

    python .\Python\tools\verify_h8_synthetic_history.py --workbook $source
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: source history" }

    python .\Python\tools\install_synthetic_recurring_ids_native.py --source $source --output $candidate
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: freeze candidate; no retry" }

    Write-Host "SYNTHETIC CC ASSIGNMENT ID NATIVE MIGRATION GREEN" -ForegroundColor Green
}
```

On ANY failure, do not delete, overwrite, retry or repair candidate A.
Never upload clinical XLSM content to GitHub or attach private data to logs.
On success, candidate A becomes frozen native evidence, not a production app.

This proves stable metadata persistence but not a permanent therapy change,
VBA Apply/Restore, production approval, or native rollback. An event store
and separate synthetic native transaction test are still required.


## Observed operator acceptance, 2026-11-12

The operator executed the prescribed native Excel block and supplied the complete success output:
- focused Python gate: `42 passed in 3.06s`;
- synthetic H8 source verification: 25 days, 522 patient-session rows,
  150 therapist-day rows, 150 attendance facts, 25 finalization events;
  source unchanged `True`;
- native installer: `SYNTHETIC NATIVE ASSIGNMENT IDS GREEN`;
  `Verified immutable assignments: 179`;
  `Original CC unchanged: YES`;
  `Historical synthetic records independently verified: YES`;
- final wrapper: `SYNTHETIC CC ASSIGNMENT ID NATIVE MIGRATION GREEN`.

**Accepted and frozen:** `UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm`.
Do not reopen for editing, overwrite, re-migrate, or use as the next change
candidate. This proves the 179 immutable IDs survived the native Excel Save
and independent readback. It does not constitute authorization to migrate
real patients or run committed effective-dated changes.


## CURRENT SAFETY NOTE (2026-11-12)

This historical native command has **already completed** on the operator's
Windows machine (see `Documentation/AUTHORITATIVE_RESUME_20261112.md`).
Its output A is accepted and FROZEN. **DO NOT RUN THIS SCRIPT AGAIN** against
the same output filename, do not overwrite/repair/delete it, and do not treat
this proof as permission to modify MASTER, CC, clinical data or permanent
therapy event history.

To resume development, start with `Documentation/ONBOARDING.md` and
`Documentation/AUTHORITATIVE_RESUME_20261112.md`. For a NEW native scenario
first implement source/tests, select a never-before-used synthetic candidate,
use one complete fail-closed PowerShell scriptblock and independently verify
history, VBA, source hashes and actual persisted semantics. On any failure
freeze that candidate and record the exact first safety error; do not retry
the same artifact.
