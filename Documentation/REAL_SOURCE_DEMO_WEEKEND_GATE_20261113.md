# Diomil real-source integrated demo: operator handoff and stop gates (2026-11-13)

## Why this exists

The operator intends to test genuine clinic scenarios over 2026-11-14/15. The priority is an **openable, integrated, private trial XLSM using the operator's reviewed real schedule**, with existing, proven application functionality. The current repository has **not** completed or natively accepted this workbook. Do not confuse synthetic CC or green CI with real-data acceptance.

This is a **read-only preflight and acceptance plan**, not a release authorization, clinical data approval, native Excel proof, or newly written importer. No patient details, IDs, paths to private files, logs containing clinical data, or workbook bytes may be committed to a public repository.

## Current source inventory verified from tracked code

- `Python/rehab_excel/trial_admission_gate.py` reads the *private review XLSX* and checks explicit PatientID confirmation, duplicate IDs, source row identity, unresolved holds, cross-specialty overlap review, valid weekday/time/specialty, mandatory providers, and shared-provider slots requiring exact approved PatientID membership.
- `Python/tools/preflight_trial_excel_sources.py` does a two-input **read-only** preflight: review XLSX + separate candidate macro-enabled intake XLSM. It records hashes before and after scanning; reports only aggregate counts/issue codes, never values; returns nonzero on blockers.
- `Python/tools/audit_pilot_source_readiness.py` inspects operational sheets, VBA, history schema/rows, registry and external links, without saving.
- `Tests/test_trial_admission_gate.py`, `Tests/test_preflight_trial_excel_sources.py` cover synthetic read-only gates, shared-slot approvals, immutable inputs, and fail-closed cases.
- `Excel/previews/UNIFIED_H10_DEMO_THERAPIST_20261110_CC.xlsm` is a **synthetic immutable source**, not an approved real-data destination; frozen A probes and migration outputs must never be used for writeback.
- Historical 61 held cells / 11 candidate overlaps: old staging observations, not current or approved patient census.

## Phase 0: privacy and operator-held prerequisites

Locate privately on the Windows PC: the **latest** `ΠΡΟΓΡΑΜΜΑ` source and reviewed staging XLSX, and identify a separate candidate XLSM explicitly intended for pilot intake. These remain institution-local. Verify which source revision produced the review. If missing/stale, do not hand-patch, assume row correspondence or import. Never enter patient names or full local private paths in chat, public GitHub, Actions, issue comments or run logs.

**Important:** the preflight's `--cc` is an *intake eligibility* input, **not permission** to modify the frozen synthetic CC file. A private, separate pilot candidate with existing operational sheets and compatible history rules must be chosen after code/source review. If no eligible candidate exists, the workflow must STOP before any write.

## Phase 1: one-block local source-only test, with all Excel windows closed

Run in PowerShell **from repository root**; this operates only on tracked code, not private data. Paste the entire scriptblock, not isolated lines. No native workbook is produced.

```powershell
& {
    $ErrorActionPreference = "Stop"
    git pull --ff-only origin work/history-statistics-rebuild
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Git pull failed" }

    python -m pytest Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: Real-source gate tests failed" }

    Write-Host "READ-ONLY REAL-SOURCE GATE TESTS GREEN" -ForegroundColor Green
}
```

These tests alone do not validate the current patient source or Microsoft Excel.

## Phase 2: local-only read-only preflight, only after operator confirms correct inputs

Edit the two **local variables**, not the code or GitHub. Do not use the frozen CC as a pilot candidate. Run the *entire* block in one go. This step does **not** produce a new demo or write/save either input. If paths are unknown or inputs are not independently reviewed, STOP instead of running.

```powershell
& {
    $ErrorActionPreference = "Stop"

    $review = "C:\PRIVATE_LOCAL_FOLDER\REVIEWED_STAGING.xlsx"
    $pilotIntake = "C:\PRIVATE_LOCAL_FOLDER\SEPARATE_PILOT_INTAKE.xlsm"

    if (!(Test-Path -LiteralPath $review -PathType Leaf)) { throw "SAFETY STOP: private reviewed source missing" }
    if (!(Test-Path -LiteralPath $pilotIntake -PathType Leaf)) { throw "SAFETY STOP: separate pilot intake missing" }

    python .\Python\tools\preflight_trial_excel_sources.py --review "$review" --cc "$pilotIntake"
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: no native real-data trial import authorized" }

    Write-Host "READ-ONLY SOURCE PREVIEW PASS: native trial still NOT performed" -ForegroundColor Green
}
```

Report **only the aggregate console summary / issue codes** after checking it contains no patient identifiers. A nonzero result is a *successful safety stop*, not an invitation to disable checks, clear held rows or mark approvals without clinical review.

## Phase 3: gate decisions before build

- **Identity:** all patient source rows have independent explicit ID confirmation; no duplicate IDs, missing identity or stale row mapping.
- **Holds and overlaps:** every affected source record is manually reconciled with provenance, or explicitly quarantined by a separate, tested scoped intake policy. **Current strict gate requires zero holds and overlaps**. It does not yet implement selective partial admission.
- **Provider/shared slots:** exact therapist/day/time and full patient membership approved in private review; approval for a subset does not cover additional patients.
- **Candidate:** identify a legitimate supported native importer into a *new* isolated workbook, schema and operational UI compatibility, and plan independent history-preservation verification. **Current read-only preflight is not an importer.**
- **Functionality:** only native-proven daily workflows go into demo acceptance checklist; permanent recurring change event-writeback, cancellation/rollback and attendance KPI without real provenance remain disabled or clearly unavailable.
- **Privacy:** only authorized staff open real-data XLSM; do not use public GitHub assets or management-share XLSX as clinical source.
- **Source integrity:** record hashes in an approved **local private** evidence file; release only aggregate safety evidence if necessary.

## Phase 4: future native trial, NOT executable from current checkpoint

After gate and importer/readback code are completed and tested, allocate one NEVER-EXISTING output path for a new private trial XLSM. Check unique output before any file operation. Persist with native Microsoft Excel on Windows, not openpyxl; save/close, reopen independently read-only, compare source hashes, verify exact patient/week-slot mapping, VBA persistence, history integrity, no phantom sessions and application navigation. On first failure, freeze the candidate, preserve the first error in a private log and use a **different output name** after fixing source; never overwrite/retry it.

Weekend operator scenarios (admission/discharge, absence/return, replacements, treatment changes, conflicts, daily close/history/reports) are a **test plan**. For each, record PASS/FAIL/NOT AVAILABLE and frozen candidate SHA privately. Do not label all capabilities working merely because the buttons exist.

## Status and handoff

As of 2026-11-13: tracked read-only admission/preflight code and synthetic tests inspected; local real input availability and freshness **not verified**, no actual current blockers recomputed, no importer built, no integrated real-data workbook produced, and no native Excel acceptance. The immediate human-dependent point is identifying and reviewing the **private** current staging plus separate eligible pilot-intake source; do not move to writeback on a guessed file. Record any real results in `CURRENT_HANDOFF.md` using only nonidentifying facts.

## Additional existing tools discovered during 2026-11-13 source inspection

- `Python/tools/prepare_trial_identity_review.py` is an **optional separate local-only preparation step** if identity mapping is not yet independently confirmed. It calls `rehab_excel.trial_identity_review.create_local_identity_review`, creates a NEW review copy rather than editing originals, and explicitly labels all inferred matches **UNCONFIRMED**. Do not execute against a destination that exists. Do not mistake suggested IDs for confirmation.
- `Python/tools/check_trial_admission.py --review ...` provides an isolated review-only count/blocker gate prior to the combined `preflight_trial_excel_sources.py`. Prefer the combined gate for a final source readiness decision. Since blocker messages may contain source row numbers, only share non-identifying aggregate output.
- `Python/tools/verify_pilot_acceptance_evidence.py --create-template ...` creates a **new** local JSON checklist, marked pending. The `--manifest ...` validation checks 13 required acceptance gates, separate SHA-256 values, pilot date and a native Excel flag. This is an evidence *structure* checker, **not** a direct test that native behavior succeeded. Keep this manifest private.
- `Python/tools/build_unified_preview.py` is a separate operational preview builder that assembles application forms and MASTER projection from an existing source XLSM, **not a reviewed staging XLSX to real-data importer**. It accepts an overwrite option; **do not use --overwrite for the real trial or on frozen evidence.**
- `Python/rehab_core/programme_mapping.py` already defines `SourceWeeklyEntry`, explicit weekday/clock mapping and patient-slot collision rejection; reuse this exact semantic model for any future real-data import plan. Do not replace it with a lossy single-clock-per-treatment mapping.

### Local acceptance manifest initialization (optional; private folder only)

This step requires no patient names, but it does require the operator's local repo and a **new, previously nonexistent** JSON destination. It does not build or authorize a demo.

```powershell
& {
    $ErrorActionPreference = "Stop"
    $evidence = "C:\PRIVATE_LOCAL_FOLDER\DIOMIL_TRIAL_EVIDENCE_20261114.json"
    if (Test-Path -LiteralPath $evidence) { throw "SAFETY STOP: evidence destination exists" }

    python .\Python\tools\verify_pilot_acceptance_evidence.py --create-template "$evidence"
    if ($LASTEXITCODE -ne 0) { throw "SAFETY STOP: evidence initialization failed" }
    Write-Host "LOCAL PENDING CHECKLIST CREATED; NATIVE PROOF NOT CLAIMED"
}
```

Do not mark any gate PASS before native evidence is actually collected.


### 2026-11-13 source-only trial import planning checkpoint
Added `Python/rehab_excel/trial_import_plan.py` (commit `b0e9d48`) and `Tests/test_trial_import_plan.py` (commit `0a9bc4c`). The plan reuses strict read-only combined preflight and exact approved `SourceWeeklyEntry` admission, cross-checks counts, binds both source SHA256 values, rejects changed sources, and hides private entries from its object repr. **It does not write an XLSM, implement a native importer, or authorize clinical use.** The new test file has six synthetic cases, but **tests/CI have NOT yet been observed passing**. Before native integration run `python -m pytest Tests/test_trial_import_plan.py Tests/test_trial_admission_gate.py Tests/test_preflight_trial_excel_sources.py -q` in one PowerShell fail-fast block; stop on red. No private patient data uploaded. Current real staging remains unverified. More work is needed to build the new isolated native importer and independent readback.
