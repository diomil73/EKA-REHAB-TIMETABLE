# Recovery Native G Build — 2026-10-31

## Source gate before build

Operator-confirmed focused MASTER/workload regression gate:

`45 passed in 4.25s`

## Fresh G unified build

Operator-confirmed canonical build of:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_G.xlsm`

completed successfully.

Builder output confirmed all canonical unified stages:

- registration menu + forms,
- patient registration bridge,
- patient edit flow,
- outpatient schedule form,
- outpatient conflict guard,
- DAILY_INPUT sheet,
- MASTER projection refresh for 93 patients,
- final user navigation,
- MASTER visual baseline.

Final builder verification reported:

- `UNIFIED PREVIEW OK`,
- `DAILY_INPUT date: 2026-10-30`,
- `Source unchanged: True`,
- VBA project present,
- DAILY_INPUT sheet present,
- application contract verified.

Independent verification reported:

- `SOURCE HASH UNCHANGED: True`,
- application contract count: `21`.

Artifact metadata reported by the operator:

- path: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_G.xlsm`,
- size: `424592` bytes,
- last write: `31/10/2026 1:07:40 πμ`,
- SHA256 begins `417CEFC6837FB9F344D186DB6F032E9416ECA646970...`.

## Safety boundary

- Production `Rehab_Center_System_v27_1.xlsm` remains immutable and its hash was unchanged.
- B, C, D, and F remain frozen evidence artifacts.
- E was never emitted after its fail-closed build failure.
- G is now the active native candidate. Do not rebuild, overwrite, or repair G in place after native testing begins.

## Required native sequence before close-day

1. Re-run the patient-absence regression with Πετιρόπουλος.
   - the affected treatment must be removed from `THERAPIST_DAILY`,
   - the corresponding treatment cell must be cleared from `MASTER_SCHEDULE`,
   - no workload button or `x/6` badges should appear on MASTER.
2. Run a therapist-absence case that produces at least one real replacement.
   - the replacement engine should run only for therapist absence,
   - the success flow should show current-action summary `patient -> replacement therapist -> replacement time`.
3. Only if both native paths are healthy may the same G candidate proceed to the first native `CloseOperationalDay` smoke.

No solver scoring, assignment rules, recurring schedule source, or close-day semantics were changed by the latest MASTER/workload fixes.
