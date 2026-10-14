# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V81.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.
- V81 MASTER toolbar/header pinning is native green.

## MASTER controls status

Native experiments V74-V79 are retired as sidebar/persistent-control attempts.

Final accepted MASTER UX:

- row 2 contains the four existing horizontal action buttons,
- row 4 contains the MASTER treatment/column headers,
- the top area is frozen through row 4,
- the application shell reapplies the pin whenever MASTER becomes active,
- no UserForm/sidebar/CommandBars mechanism is used.

Implementation commit: `0c4b25eac6378e97366b6690b3443c5d46948f53`

Regression commits:

- `2860fe43bd133e2e2bac663b422ff54f5985f2c4`
- `b05426cd9cf68b39267159d84ce3d1982c145af8`

Native V81 verification: green. Vertical scrolling keeps both the four MASTER buttons and row-4 treatment headers visible.

## Patient-program native verification

Confirmed green by the user:

1. edit-existing-program end to end,
2. existing program preload,
3. atomic resave without false self-conflict,
4. optional manual names persistence/reload for Ρομποτικό / Εργο / Λογο / Ψυχ.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- cleanup/reset focused gate: `42 passed in 3.70s`
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- V81 native toolbar/header pinning gate: green
- optional manual names persistence/reload: native green
- **full regression before cleanup: `722 passed, 19 warnings`**

## Current focus: retired single-treatment scheduling cleanup

The full regression is green, so cleanup has started from a known-good baseline.

Removed/retired source:

- `Python/rehab_excel/patient_scheduling_flow.py`
- `Python/rehab_excel/patient_scheduling_form_vba.py`
- obsolete `plan`, `validate`, `plan-vba`, `validate-vba`, `recurring-vba`, and `validate-recurring-vba` commands from `Python/tools/patient_scheduling_bridge.py`

Retained current behavior:

- `auto-propose` / `auto-propose-vba` in `patient_scheduling_bridge.py`, used by the joint patient-program auto scheduler,
- `PatientProgramRuntime` / current `frmPatientScheduling`,
- legacy VBA component removal/verification in `patient_registration_scheduling_slice.py` so old workbooks cannot retain `PatientSchedulingRuntime`.

Cleanup commits so far:

- `93e2759726f002dd86c82168e6d3abfcf40a5d32` bridge cleanup
- `ccc97c6f9e43cf32ca085cb42ba03f12469acb50` remove retired scheduling flow
- `8c0bdd326e6727c04594a6284e0148f6846bdf74` remove retired scheduling form source
- `8265d4000036a15323d0a0cb4ec95d6cc340e0f3` remove retired flow tests
- `29a9221055f2057f9062af7f8853274271724ef2` remove retired form tests
- `b0d40de9f3e0eb1b76ce9cca5b1a36aa1bda3c5f` keep bridge regression tests on joint auto path

## Immediate next verification gate

1. Pull latest branch.
2. Run `python -m pytest`.
3. If green, cleanup is accepted.
4. Then update final project status / roadmap and decide the next roadmap milestone.

## Business rules that must not regress

- ΦΘ provider is required and solver-owned.
- Ρομποτικό / Εργο / Λογο / Ψυχ provider/name is optional free text and is not a solver constraint.
- Ανακλινόμενο / Πισίνα / ΕΦΑ have no provider/name field.
- Nonblank manual constraints are never silently overwritten.
- One final atomic save writes the included patient program.
- Standard treatment duration: 45 minutes.
- Ανακλινόμενο duration: 30 minutes.
- Exact adjacency is legal.
- `PATIENT_PLANNER` authoritative, MASTER derived.
- Conflict override must be explicit; return-for-correction performs zero writes.

## Recovery order if context is lost

1. `Documentation/CURRENT_HANDOFF.md`
2. `Documentation/IMPLEMENTATION_MAP.md`
3. `Documentation/PROJECT_STATUS.md`
4. `Documentation/DECISIONS.md`
5. active tests

Windows: always `python -m pytest`, never bare `pytest`.
