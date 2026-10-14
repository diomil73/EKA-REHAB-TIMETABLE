# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V80.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.

## MASTER persistent-controls conclusion

Native experiments V74-V79 are retired:

- V74/V75: freeze-based pinning did not keep the worksheet buttons visible.
- V76: legacy `Application.CommandBars` sidebar did not produce reliable visible UI.
- V77: modeless UserForm had not reached the final workbook because the preview builder kept stale shell VBA.
- V78 proved the final shell-refresh build fix works and the UserForm can remain visible during scroll.
- V79 showed the modeless UserForm is not operationally acceptable: inconsistent appearance, Alt+F8/Alt+F4 interaction, and unstable window behavior.

Decision: stop workbook-only sidebar experiments. Keep the normal top MASTER toolbar for the current `.xlsm` milestone. A true docked sidebar would require a separate Excel add-in/custom task-pane architecture.

V80 native reset verification is green:

- no sidebar/UserForm appears,
- no strange freeze/split state remains,
- normal top MASTER toolbar works,
- navigation works,
- Save & Exit works.

## Cleanup retained

`Python/rehab_excel/application_shell_vba.py`:

- has no `ShowMasterSidebar` / `HideMasterSidebar` runtime logic,
- has no runtime freeze/split pinning,
- removes retired `frmMasterSidebar` if found in older workbooks,
- keeps the stable application chrome, MASTER activation/view, navigation sidecar, and Excel restore behavior.

Important build-pipeline fix retained:

- the final patient-scheduling preview reinstalls the current `modApplicationShell`,
- verifies the shell entry points,
- verifies the retired sidebar form is absent before saving.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- persistent-controls history gates: `39`, `39`, `41`, `42`, `44`, `45` passes at successive iterations
- cleanup/reset focused gate: `42 passed in 3.70s`
- V80 native reset gate: green

## Current focus: edit-existing-program end to end

The persistent-controls detour is closed. Resume the patient-program milestone from the existing edit flow.

Native verification target:

1. MASTER / Βασικό μενού → **Επεξεργασία ασθενή**.
2. Select an existing patient with an existing patient program.
3. Save any registry edits.
4. Confirm the same patient-program form opens.
5. Confirm current treatments/day/time/provider values preload correctly.
6. Modify one or more treatment assignments.
7. Save atomically.
8. Verify `PATIENT_PLANNER` and MASTER composite rendering reflect the edited program.
9. Confirm no duplicate/self-conflict is created from the patient's old persisted assignments.

After that:

- optional manual names persistence/reload for Ρομποτικό / Εργο / Λογο / Ψυχ,
- full regression suite,
- retired single-treatment cleanup.

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
