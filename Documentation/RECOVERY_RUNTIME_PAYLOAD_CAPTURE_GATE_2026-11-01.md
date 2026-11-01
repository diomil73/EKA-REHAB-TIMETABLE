# Recovery runtime payload capture gate — 2026-11-01

## Why this gate exists

Native U failure evidence and fresh manual payload evidence disagree on patch ordering:

- native U log at 2026-11-01 12:37:33 reports `patch_index=2`, `MASTER_SCHEDULE!D8`,
- a fresh payload prepared from the same U workbook reports `MASTER_SCHEDULE!D8` as raw `PATCH #11` and `MASTER_SCHEDULE!M8` as `PATCH #12`.

This discrepancy had to be resolved before changing verifier semantics again.

## What was ruled out

### Wrong workbook VBA / wrong script path

`inspect_u_runtime_provenance.py` proved:

- U contains the character-strike helper,
- U contains the bounded tri-state non-active overlay branch,
- U contains the whole-cell-first verifier helper,
- `ResolveDailyInputPreparationScript()` resolves to the repository `Python/tools/prepare_daily_input_in_place.py`,
- resolved script SHA256 equals repository script SHA256.

### Late addition of `rebuild_multi_member_groups=True`

`git blame` proves the flag exists from the initial in-place preparer commit:

`17d55687 (2026-10-30 23:29:40 +0200) rebuild_multi_member_groups=True`

The original file history also shows this flag was present when `prepare_daily_input_in_place.py` was first introduced, so it is not a later regression.

### Nondeterministic patient-centric render ordering

`patient_centric_preview.py` emits patches from:

`sorted(renders.items(), key=lambda item: (item[0][1], item[0][0].casefold()))`

so Python hash/set ordering is not an explanation.

### Recent drift in the patient-centric renderer

The renderer history is old and stable relative to the recovery work. Important commits are from 19–20 September 2026, including `9f0a114`, `e47c6d4`, and `d2f09cd`.

### CHANGE rows causing the patch-number mismatch

The ledger-aware apply path counts only rows whose kind is `PATCH` when it increments `currentPatchIndex`. `CHANGE` rows do not explain native `D8 = patch #2` versus fresh `D8 = patch #11`.

### Patient-centric/business engine failure

The copy-based `Python/tools/apply_daily_input.py` path successfully applied the same U patient absence to a fresh diagnostic workbook.

Observed proof:

- `DAILY INPUT APPLIED PREVIEW OK`
- `Changed sessions: 1`
- affected patient: `planner:4:ΦΘ@2026-10-30 [B7]`
- `Input unchanged: True`
- `VBA preserved: True`
- source SHA256 unchanged before/after.

Therefore the business/planning engine remains functional and must not be rewritten based on the earlier native rollback symptom.

## Diagnostic source change

Added deterministic source migrator:

`Python/tools/apply_daily_input_runtime_capture_source_fix.py`

On an in-place apply failure it changes only the failure path so the exact temp artifacts are copied before normal cleanup:

- `%TEMP%\eka_daily_input_last_payload.tsv`
- `%TEMP%\eka_daily_input_last_response.json`

The existing error log remains:

- `%TEMP%\eka_daily_input_last_error.log`

Business logic, patch construction, patch application, rollback, verifier acceptance rules, and success-path behavior are unchanged.

Regression tests:

`Tests/test_daily_input_runtime_capture_source_fix.py`

The migrator fails closed unless the expected source marker exists exactly once.

## Green source gate

User ran the focused regression suite after applying the source migrator.

Result:

`14 passed in 4.00s`

Accepted canonical source commit:

`1b039de796a6609a3bf16983b7b8beb67bfc0685` — `Capture failing DAILY_INPUT runtime payload`

## V native provenance

Fresh workbook:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261101_V.xlsm`

Provenance inspection showed:

- workbook SHA256: `1dd912b91d232e631897fa99bf2a309d39c809add5fa8d0b1b912844695ec80e`,
- `MODULE_LINES: 660`,
- character-strike helper present,
- tri-state non-active branch present,
- whole-cell-first verifier comment present,
- resolver points to repository `Python/tools/prepare_daily_input_in_place.py`,
- resolved script hash equals repository script hash,
- branch HEAD was `1b039de` when the native V was built/run.

Therefore V definitely contained the intended runtime-capture source.

## V native outcome — FIRST GREEN IN-PLACE PATIENT ABSENCE

The runtime capture files did not exist after the V run. Read-only post-run inspection proved this was because the validated apply did **not fail**. The patient absence was durably committed.

Observed V state:

### DAILY_INPUT

- date: `30/10/2026`,
- row 21: patient `ΖΑΛΟΚΩΣΤΑΣ`, all day `ΝΑΙ`, status/reason `διακομισμένος`.

### THERAPIST_DAILY

`B7`:

- value: `ΖΑΛΟΚΩΣΤΑΣ`,
- `strike=True`,
- muted/gray font preserved,
- infectious yellow fill preserved.

Other rebuilt multi-member cells were also present and readable, including D2, H2, I13, J13, E14, E15, G6, J16, F17.

### MASTER_SCHEDULE

`D8`:

- original formula preserved,
- `strike=True`,
- font color red,
- fill preserved.

`M8`:

- value: `διακομισμένος`,
- font color red,
- no whole-cell strike.

Nearby unaffected row 9 remained normal, confirming the overlay did not smear into adjacent patient rows.

### DAILY_CHANGE_LEDGER

Row 2 was persisted:

- ChangeID: `daily-input:2026-10-30:patient:21`,
- ChangeType: `patient_absence`,
- Subject: `ΖΑΛΟΚΩΣΤΑΣ`,
- FromText: `Κανονικό πρόγραμμα`,
- ToText: `Απουσία ασθενούς όλη ημέρα`,
- Reason: `διακομισμένος`,
- SourceRow: `21`,
- Status: `ACTIVE`.

## Architectural conclusion after V

The earlier native blocker `DAILY_INPUT read-back mismatch` at `MASTER_SCHEDULE D8/D9` is now cleared for this patient-absence scenario.

This is the first native proof that the current in-place transaction can:

1. read DAILY_INPUT,
2. compute the patient-centric operational plan,
3. apply THERAPIST_DAILY changes,
4. apply MASTER presentation/status overlay,
5. pass semantic read-back verification,
6. persist an ACTIVE daily-change ledger row,
7. save the same workbook without rollback.

Therefore:

- do not reopen the D8/D9 verifier investigation unless new contradictory evidence appears,
- do not modify replacement/global solver,
- do not modify patient-centric renderer,
- do not interpret missing failure-capture files as a problem when the transaction succeeded.

## Next gate

The recovery blocker now advances to **native restore proof**.

Use one fresh workbook for an end-to-end absence + restore scenario. Do not reuse V for a new native test.

Required proof:

1. apply the same patient absence successfully,
2. confirm an ACTIVE ledger row exists,
3. restore that change through the workbook UI,
4. verify DAILY_INPUT row is restored/cleared as designed,
5. verify THERAPIST_DAILY B7 returns to active presentation,
6. verify MASTER D8/M8 return to normal projection/presentation,
7. verify ledger row becomes `RESTORED` with `RestoredAt`,
8. verify no unrelated THERAPIST_DAILY content is erased,
9. verify same-workbook save succeeds.

Only after native restore is green may the close-day/history track resume.

## Do not do

- Do not reuse U or V for another native repair experiment.
- Do not repair-over-repair a failed workbook.
- Do not modify replacement/global solver.
- Do not modify patient-centric renderer.
- Do not modify MASTER verifier semantics without new failure evidence.
- Do not revive N/O/P/Q post-build recovery patching.
- Do not touch the production baseline workbook.

This checkpoint is the authoritative resume point after the first green native in-place patient absence.
