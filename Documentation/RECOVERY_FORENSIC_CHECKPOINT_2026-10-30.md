# Recovery Forensic Checkpoint — 2026-10-30

## Purpose

This checkpoint records the forensic recovery work completed on 2026-10-30 before any further implementation changes. It exists to prevent a future conversation or handoff from repeating the same workbook/source confusion.

No operational source code was changed as part of this checkpoint. No production workbook was modified.

## Hard safety rules

1. `Rehab_Center_System_v27_1.xlsm` remains immutable.
2. `baseline/known-good-4a2eed0` is the exact immutable trusted baseline pointer and must not be used for experiments.
3. `recovery/known-good-4a2eed0` is not to be assumed immutable because it later advanced remotely.
4. New implementation work belongs on `work/history-statistics-rebuild` or a child work branch.
5. Existing workbook artifacts are forensic references only. Never overwrite them.
6. Every native Excel test must create a newly named `.xlsm` copy.
7. Do not use `git reset --hard`, `git clean`, or destructive workbook repair as recovery tools.
8. Do not treat a subsystem smoke workbook as a complete application artifact unless the full application contract passes.

## Local Git safety result

After `git fetch origin`, the command:

```powershell
git log --branches --not --remotes --oneline --decorate
```

returned no commits. Therefore there were no commits reachable from current local branches that were absent from all remotes at the time of the check.

This does **not** mean every file on disk exists on GitHub. Several `.xlsm` artifacts are local/untracked.

## Recovered local workbook lineage

A large local workbook lineage was found under `Excel/previews`.

Key recovered sequence:

- `MASTER_TOOLBAR_SMOKE_V81.xlsm`
- `DAILY_CHANGE_LEDGER_V82.xlsm`
- `DAILY_CHANGE_LEDGER_V83.xlsm`
- `DAILY_CHANGE_LEDGER_V84.xlsm`
- `DAILY_CHANGE_LEDGER_V85.xlsm`
- `DAILY_CHANGE_LEDGER_V86.xlsm`
- `ORPHAN_SMOKE_INPUT_V86.xlsm`
- `ORPHAN_RESOLUTION_V87.xlsm`
- `ORPHAN_RESOLUTION_V88.xlsm`
- `AUTHORITATIVE_RESTORE_V89.xlsm`
- `RERESOLVED_V90.xlsm`
- `SELECTED_RESTORE_V91.xlsm`
- `RESTORE_ALL_V92.xlsm`

The lineage then continues through outpatient, replacement, double-booking, close-day and DAILY_INPUT in-place smoke artifacts.

The user's recollection of a V81 → V86 progression was therefore confirmed by the local filesystem.

## Forensic copies created

Copies were created under:

`Forensic_Backups/Recovery_Lineage`

The following forensic copies were hashed with SHA256:

- `MASTER_TOOLBAR_SMOKE_V81.xlsm` — `912CCA71614196572FD4E3FCED195942B19A72CBB519D590A38D25F5D5011850`
- `DAILY_CHANGE_LEDGER_V82.xlsm` — `637E69D41A380138FE81E89DDA7690ED2C0F569E8C6D97BBFD24B18AC907CF29`
- `DAILY_CHANGE_LEDGER_V83.xlsm` — `A1286A3C7ED4B3F2C84847CAE33FD12DAAF64F41598C12DAC56B301C89D49692`
- `DAILY_CHANGE_LEDGER_V84.xlsm` — `6F8247F30B042BAB47656019FC9B96BD5420AA0B8E94CDF3FFF7B70E0F165AB8`
- `DAILY_CHANGE_LEDGER_V85.xlsm` — `34C814B036FAD7A3715F36FA8B1206C2E5FE6FD24C72A39008E1EFFC6D8608DE`
- `DAILY_CHANGE_LEDGER_V86.xlsm` — `A0FD0ED96DD58C1AB06FEFA26675EABEC7A13DDFB641238188B7CFC0E85F52FD`
- `ORPHAN_SMOKE_INPUT_V86.xlsm` — `BA732BB50DA024C39DF0D7C91FAE72EB00FE3B03A1409A3503F517426BBEE7FC`
- `ORPHAN_RESOLUTION_V87.xlsm` — `F26994791253C9F6D4E03A38AF3436301569FC2BC9B1D6BD0967AEE4A5211BAA`
- `ORPHAN_RESOLUTION_V88.xlsm` — `87E2B80EC7D4F98217970D35F73A14D47ACAE7395B36B22D7CFC7505283BA94F`
- `AUTHORITATIVE_RESTORE_V89.xlsm` — `7B9F0D4EB99F7F6EC8BF236AD28BE7657AAD4E3C2509EB913C9EC93E2F1E5D1F`
- `RERESOLVED_V90.xlsm` — `9777006BCEF2508CC9E1740DE1B1EBA49AFB2FC6926FE43B5F46CF64291C5194`
- `SELECTED_RESTORE_V91.xlsm` — `30B769201111BE95F7761266EBDC61264155D37FD1CDC1FFC31C779B7A0333ED`
- `RESTORE_ALL_V92.xlsm` — `E753FDCE0FA857165EBF96CC3F8CDA10881D4CAABB73E1106337C88F2685C335`
- `OUTPATIENT_RECURRING_SMOKE.xlsm` — `E79FBFBBE7A287BCC24EE9D6BE81393D30C907FA45B37974EDCFC089D9DB1F2A`
- `CLOSE_DAY_VERIFIED_PREVIEW.xlsm` — `73C0568E8D5E17C6D22D8B7B41A722DCACE291221455B924240ABA41164E5200`
- `DAILY_INPUT_IN_PLACE_V16_OPERATOR.xlsm` — `C0BD7233AFD40E2D677DE6DBCE55D4C5234E925C80D5AE0016AA56371CF7FAB3`
- `DAILY_INPUT_RESTORE_SMOKE.xlsm` — `214B0C72349987D3D00F2C723FFADE3346FD714995DEADC8AEDC04DA3F18CF9C`

These hashes are the immutable forensic fingerprints for future comparison.

## PATIENT_SCHEDULING_PREVIEW finding

Two local/untracked workbooks were found:

- root `PATIENT_SCHEDULING_PREVIEW.xlsm`
- `Preview/PATIENT_SCHEDULING_PREVIEW.xlsm`

They are different files by SHA256 but have the same sheets, defined names, VBA component inventory and extracted VBA source code. Their `workbook.xml` difference is only the saved absolute path. Their `vbaProject.bin` files differ internally, but the extracted VBA source is identical.

Conclusion: these two files are useful forensic references but are **not** two different source-code generations and are not the main recovery baseline.

## Workbook lineage findings

### V81 → V86

`MASTER_TOOLBAR_SMOKE_V81` to `DAILY_CHANGE_LEDGER_V86` is additive:

- no modules removed
- `modDailyChangeLedger` added
- sheet module `Φύλλο17` added

### V86 → V92

`DAILY_CHANGE_LEDGER_V86` to `RESTORE_ALL_V92` is also additive:

- no modules removed
- sheet modules `Φύλλο18` and `Φύλλο19` added
- workbook now contains `DAILY_REPLACEMENT_FACTS` and `DAILY_STATISTICS`

### V92 → OUTPATIENT_RECURRING_SMOKE

This is a substantial outpatient evolution:

Added:

- `frmOutpatientScheduleManager`
- `modOutpatientScheduleEdit`
- `modOutpatientScheduleManager`
- `modTherapistWorkloadStrip`

Multiple existing modules were also changed, including `frmOutpatientSchedule`, `modOutpatientInSession`, registration UI, conflict handling, patient scheduling and ledger code.

This artifact is a strong outpatient/application reference.

### OUTPATIENT_RECURRING_SMOKE → DAILY_INPUT_IN_PLACE_V16_OPERATOR

The later DAILY_INPUT/operator smoke adds important operational components:

- `frmDoubleBookingReview`
- `frmReplacementDecision`
- `modCloseDayAction`
- `modDailyInputChangeLedger`
- `modDailyInputChangeNavigation`
- `modDoubleBookingReview`
- `modGlobalReplacementAction`
- `modReplacementOperator`
- `modSupervisorPanelStyle`

It also changes many existing modules.

However it removes:

- `frmOutpatientSchedule`
- `modOutpatientSchedule`

Therefore `DAILY_INPUT_IN_PLACE_V16_OPERATOR.xlsm` must **not** be treated as a complete application artifact by itself.

## Outpatient form was not intentionally retired

A strict procedure-definition comparison showed that many old `frmOutpatientSchedule` procedures are absent from the V16 operator smoke, including the suggestion UI, JSON request builders, bridge resolution, patient selection and edit-target helpers.

Repository source confirms that the outpatient form remains part of the intended architecture:

- `outpatient_schedule_form_vba.py` still defines `frmOutpatientSchedule` and `modOutpatientSchedule`
- `outpatient_schedule_edit_vba.py` loads and uses the form
- `outpatient_conflict_guard_vba.py` loads the form and calls `SelectPatientById`
- `patient_edit_flow_vba.py` opens the form
- `registration_menu_vba.py` opens the form
- `outpatient_schedule_manager_vba.py` still references `ShowOutpatientScheduleForm`
- the Python bridge and suggestion bridge tools still exist

Conclusion: the V16 operator workbook is a subsystem smoke, not proof of a deliberate outpatient-form replacement.

## Canonical builder finding

`Python/tools/build_unified_preview.py` is the intended full sequential builder. Its steps include:

1. registration menu + forms
2. patient registration bridge
3. patient edit flow
4. outpatient schedule form
5. outpatient conflict guard
6. DAILY_INPUT sheet
7. MASTER projection refresh
8. final navigation
9. MASTER visual baseline

This means the intended build path explicitly re-installs `frmOutpatientSchedule` after the registration-menu preparation stage.

`build_registration_menu_preview.py` performs a clean-copy rebuild. It removes old operational/app components from a temporary copy and then installs several operational modules including global replacement, replacement operator, double-booking review, workload, daily change ledger, replacement ledger sync and DAILY_INPUT change navigation.

Therefore a registration-menu preview or another subsystem smoke must never again be assumed to be the complete application merely because it contains many operational modules.

## Application contract gap

`Python/rehab_excel/application_contract.py` currently requires only a subset of the true application stack. Current required components include the shell, toolbar, registration menu/forms, DAILY_INPUT action, workload strip and outpatient manager.

The contract does **not** currently require all critical components such as:

- `modOutpatientSchedule`
- `frmOutpatientSchedule`
- close-day action
- replacement operator module/form
- global replacement action
- double-booking review module/form
- daily change ledger / DAILY_INPUT change ledger/navigation
- other operational modules that the full app depends on

This is a systemic safety gap because `assert_application_contract()` is used as an input/output gate by multiple important tools, including confirmed double booking, global replacement, double-booking review preparation/installation, DAILY_INPUT preview, multi-absence smoke preparation and replacement-fact restore.

A workbook can therefore currently pass the application contract while still being incomplete for real application use.

## Close-day installation question still open

`build_registration_menu_preview.py` imports `CLOSE_DAY_MODULE_NAME` and removes any old close-day module during pre-clean. In the inspected `_install_operational_modules()` body, close-day installation was not visibly present.

Before any new build is run, verify exactly where `modCloseDayAction` is installed in the canonical pipeline. Do not assume it is present.

## Correct recovery interpretation

The correct strategy is **not** to choose one old workbook and continue mutating it.

The recovered artifacts show two important capability lines:

- outpatient/create/edit/conflict/workload capability
- DAILY_INPUT/replacement/double-booking/close-day/operator capability

The source tree already contains a canonical pipeline intended to assemble the application. Recovery work must strengthen that pipeline and its contract rather than manually merging binary workbook artifacts.

## Exact resume point for next session

Do not build anything before these checks are complete.

1. Search all Python source for `install_close_day`, `modCloseDayAction`, and `CLOSE_DAY_MODULE_NAME` to establish the authoritative close-day install path.
2. Define the complete required application component set from the recovered known-good capability union.
3. Add focused tests that fail when any critical component is missing.
4. Strengthen `application_contract.py` only after the required set is proven.
5. Verify `build_unified_preview.py` / `build_registration_menu_preview.py` install every required component exactly once and in the correct order.
6. Run Python regression gates before any Excel mutation.
7. Build a **fresh uniquely named** unified smoke workbook from the immutable source/canonical pipeline.
8. Verify its VBA inventory and required sheets automatically.
9. Only then perform native Excel interaction tests.
10. Resume statistics/history implementation only after the complete application build path is proven safe, unless a separate pure-Python slice can proceed without workbook integration.

## Do not repeat

- Do not use `REGISTRATION_MENU_PREVIEW.xlsm` as the complete app.
- Do not use `DAILY_INPUT_IN_PLACE_V16_OPERATOR.xlsm` as the complete app.
- Do not patch the forensic V81–V92 artifacts.
- Do not repair over a failed smoke.
- Do not overwrite known-good previews.
- Do not reopen the replacement solver merely because an application assembly artifact is incomplete.
- Do not infer that a newer smoke is more complete than an older one without checking its component contract.

## Documentation follow-up

On the next implementation session, after the full application contract is corrected and tested, update `CURRENT_HANDOFF.md`, `MASTER_REQUIREMENTS_RECOVERY_PLAN.md`, `ONBOARDING.md` where the project-wide mental model changes, and the relevant specialist architecture/checkpoint document.
