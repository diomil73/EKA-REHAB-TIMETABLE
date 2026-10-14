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

Native experiments V74-V79 are retired as sidebar/persistent-control attempts:

- V74/V75: earlier freeze approaches did not keep the intended UI reliably visible.
- V76: legacy `Application.CommandBars` sidebar failed.
- V77-V79: modeless UserForm path was ultimately rejected as operationally unreliable.
- V80 restored a clean native-green MASTER baseline.

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

Native V81 verification: green. Vertical scrolling keeps both the four MASTER buttons and row-4 treatment headers visible, with no unstable sidebar/window behavior.

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- persistent-controls history gates: `39`, `39`, `41`, `42`, `44`, `45` passes at successive iterations
- cleanup/reset focused gate: `42 passed in 3.70s`
- V80 native reset gate: green
- row-4 toolbar/header pinning focused source gate: `43 passed in 1.92s`
- V81 native toolbar/header pinning gate: green
- optional manual names persistence/reload: native green

## Current focus

The MASTER UX milestone is closed.

Patient-program native checks now confirmed green:

1. edit-existing-program end to end,
2. existing program preload,
3. atomic resave without false self-conflict,
4. optional manual names persistence/reload for Ρομποτικό / Εργο / Λογο / Ψυχ.

Immediate next step:

1. Run the full regression suite with `python -m pytest`.
2. If green, update the implementation map/project status.
3. Then clean up retired single-treatment scheduling code.

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
