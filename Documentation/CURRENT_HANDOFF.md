# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V79.xlsm`
- V73 scheduling/conflict behavior remains native green.

## MASTER persistent-controls conclusion

Native experiments V74-V79 are retired:

- V74/V75: freeze-based pinning did not keep the worksheet buttons visible.
- V76: legacy `Application.CommandBars` sidebar did not produce reliable visible UI.
- V77: modeless UserForm had not reached the final workbook because the preview builder kept stale shell VBA.
- V78 proved the final shell-refresh build fix works and the UserForm can remain visible during scroll.
- V79 showed the modeless UserForm is not operationally acceptable: inconsistent appearance, Alt+F8/Alt+F4 interaction, and unstable window behavior.

Decision: stop workbook-only sidebar experiments. Restore the stable application shell and keep the normal top MASTER toolbar for now. A true docked sidebar would require a separate Excel add-in/custom task-pane architecture and is not part of the current `.xlsm` milestone.

## Cleanup implemented

`Python/rehab_excel/application_shell_vba.py` now:

- removes all `ShowMasterSidebar` / `HideMasterSidebar` runtime logic,
- contains no `FreezePanes`, `SplitRow`, `SplitColumn`, `vbModeless`, or CommandBar sidebar code,
- explicitly removes retired `frmMasterSidebar` if it is present in an older workbook,
- keeps the normal application chrome, MASTER activation/view, navigation sidecar, and Excel restore behavior.

Important build-pipeline fix retained:

- the final patient-scheduling preview still reinstalls the current `modApplicationShell`,
- verifies `EnterApplicationShell` and `KeepApplicationShell`,
- verifies the retired `frmMasterSidebar` is absent before saving.

Cleanup commits:

- `a0158a479904c26cff8fe1f4265ad777dd8426e6` application shell reset
- `c7760808f2f5ad5b5f6798a7e34521ab0e50f3f5` retired-sidebar regression tests
- `87e6388cd73c1f62f67761d96130d1379c0904a4` baseline shell assertions
- `effadb2973326046b0c395c72f71916647549c4e` final shell refresh without sidebar
- `a2fb0c266975c7322fe93f3733043a2a4807b2d2` preview-shell cleanup regression

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- persistent-controls history gates: `39`, `39`, `41`, `42`, `44`, `45` passes at successive iterations
- latest pre-cleanup focused gate: `45 passed in 1.34s`
- cleanup/reset focused gate: `42 passed in 3.70s`

## Immediate next start point

1. Build one clean reset preview from V48.
2. Verify in native Excel:
   - no sidebar/UserForm appears,
   - no strange split/freeze state remains,
   - normal top MASTER toolbar works,
   - navigation and Save & Exit still work.
3. Then resume the patient-program milestone:
   - edit-existing-program end to end,
   - optional manual names persist/reload,
   - full regression suite.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal.
- MASTER is derived; `PATIENT_PLANNER` remains authoritative.
- Conflict override must be explicit; return-for-correction performs zero writes.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
