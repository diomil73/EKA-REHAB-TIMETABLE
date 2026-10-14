# Current Handoff

Last updated: 2026-10-14

## Branch / working line

- Branch: `feature/outpatient-scheduling`
- Clean preview base: `Excel/previews/MASTER_TOOLBAR_SMOKE_V48.xlsm`
- Latest native preview tested: `MASTER_TOOLBAR_SMOKE_V80.xlsm`
- V73 scheduling/conflict behavior remains native green.
- V80 clean reset preview is native green.

## MASTER controls status

Native experiments V74-V79 are retired as sidebar/persistent-control attempts:

- V74/V75: earlier freeze approaches did not keep the intended UI reliably visible.
- V76: legacy `Application.CommandBars` sidebar failed.
- V77-V79: modeless UserForm path was ultimately rejected as operationally unreliable.
- V80 restored a clean native-green MASTER baseline.

New clarified requirement from the user: while scrolling vertically, at minimum the treatment column headers must remain visible, and the four MASTER action buttons must also follow either horizontally or vertically.

Chosen implementation: keep the existing horizontal top toolbar and freeze the MASTER top area through row 4.

This is intentionally narrower than the retired sidebar experiments:

- row 2 contains the four existing action buttons,
- row 4 contains the MASTER column/treatment headers,
- the application shell reapplies the pin whenever MASTER becomes active,
- no UserForm/sidebar/CommandBars mechanism is used.

Implementation commit: `0c4b25eac6378e97366b6690b3443c5d46948f53`

Regression commits:

- `2860fe43bd133e2e2bac663b422ff54f5985f2c4`
- `b05426cd9cf68b39267159d84ce3d1982c145af8`

## Latest confirmed test checkpoints

- narrow automatic-scheduling gate: `16 passed in 2.87s`
- wider patient-program integration gate: `71 passed in 3.66s`
- persistent-controls history gates: `39`, `39`, `41`, `42`, `44`, `45` passes at successive iterations
- cleanup/reset focused gate: `42 passed in 3.70s`
- V80 native reset gate: green
- new row-4 toolbar/header pinning: source changes ready, focused pytest pending

## Current focus

1. Run focused application-shell / toolbar / preview-shell-refresh tests.
2. If green, build the next preview from clean V48.
3. Native verify that vertical scroll keeps both:
   - the four horizontal MASTER buttons,
   - row-4 treatment headers.
4. Confirm no sidebar/UserForm appears and no odd split state is introduced.
5. Then continue patient-program verification. The edit-existing-program path was already reported native green by the user.
6. Next patient-program checks: optional manual names persistence/reload, then full regression.

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
