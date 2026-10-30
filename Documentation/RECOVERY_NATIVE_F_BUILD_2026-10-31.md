# Recovery Native F Build — 2026-10-31

## Source gate before build

Operator-confirmed focused COM-safe operator-visibility gate:

`34 passed in 2.58s`

## Fresh F unified build

Operator-confirmed canonical build of:

`Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_F.xlsm`

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

- path: `Excel/previews/UNIFIED_CLOSE_DAY_RECOVERY_20261031_F.xlsm`,
- size: `428556` bytes,
- last write: `31/10/2026 12:42:07 πμ`,
- SHA256 begins `3911C155C1BDF2A892FF6D982DD87130C325F8C9C53...`.

## Safety boundary

- Production `Rehab_Center_System_v27_1.xlsm` remains immutable and its hash was unchanged.
- B, C, and D remain frozen evidence artifacts.
- E was never emitted after its fail-closed build failure and its name is not reused.
- F is frozen as native evidence after the findings below. Do not rebuild, overwrite, or repair F in place.
- Close-day native testing remains blocked until a fresh artifact verifies the patient-absence MASTER projection fix and the corrected workload UI placement.

## Native F findings

Operator-confirmed behavior on F:

1. Normal DAILY_INPUT no longer auto-renders therapist `x/6` badges on MASTER.
2. The recovery-only `Φόρτος θεραπευτών` button still remained on MASTER and displayed workload when clicked.
3. This button placement is not the agreed final UX. The authoritative UI roadmap places current workload and workload history in the therapist personal card opened from the therapist identity, not as MASTER badges/button.
4. A patient absence was then tested with patient Πετιρόπουλος.
5. The patient was correctly removed from the daily therapist/operational schedule.
6. The same patient treatment remained visible on `MASTER_SCHEDULE`, creating a cross-view inconsistency.
7. No replacement flow was expected for this patient absence. Replacement resolution is the therapist-absence path.

## Confirmed MASTER projection cause

`master_daily_overlay.py` only emitted MASTER patches for `DailySessionStatus.REPLACED` and skipped all other daily statuses. Therefore a `PATIENT_ABSENT` or explicit `CANCELLED` state could be removed from the daily therapist projection while the recurring MASTER treatment cell remained unchanged.

Bounded source fixes:

- `7f4884c` extends the MASTER daily overlay to emit `WriteIntent.CLEAR` for inpatient `PATIENT_ABSENT` and `CANCELLED` states, scoped only to the affected treatment cell,
- pending `THERAPIST_ABSENT` is deliberately not cleared before replacement resolution,
- recurring/base source rows remain untouched,
- `fd1afbb` adds regressions for replacement overlay, patient-absence clear, explicit-cancellation clear, pending therapist absence, precomputed projection, and outpatient exclusion.

`prepare_daily_input_in_place.py` already composes `build_master_daily_overlay_plan` into the validated in-session payload, so the new clear patch follows the established native apply/verify/rollback path.

## Corrected workload UI direction

The authoritative UI roadmap specifies therapist current workload/history inside the therapist personal card, reached from therapist identity in the application. MASTER is not the workload dashboard.

Bounded source fixes:

- `f0cfe7b` converts `modTherapistWorkloadStrip` from MASTER presentation UI into a backend workload helper exposing current workload and capacity,
- canonical installation removes legacy `ekaWorkload_*` shapes from MASTER and no longer creates a workload button or badges,
- `5bb950a` locks the backend-only workload contract and legacy-shape cleanup in regression tests.

The therapist personal card itself remains a later UI/reporting slice; this recovery change only removes the incorrect MASTER presentation while preserving workload calculation for that future card.

## Post-fix focused gate

Operator-confirmed focused regression gate after the MASTER-clear and workload-placement fixes:

`45 passed in 4.25s`

Covered:

- MASTER daily overlay,
- in-session DAILY_INPUT preparation/action,
- therapist workload backend-only contract,
- canonical workload installation/legacy-shape cleanup,
- operator visibility patch,
- production actions stay in-session,
- unified builder regressions.

This source gate is green for the fresh G build. F remains frozen and must not be modified.

## Next exact step

1. Build a fresh uniquely named `UNIFIED_CLOSE_DAY_RECOVERY_20261031_G.xlsm` from immutable `Rehab_Center_System_v27_1.xlsm`; do not reuse F.
2. Verify source SHA unchanged and application contract remains green.
3. Native-test the same Πετιρόπουλος patient absence and confirm both the daily projection and the matching MASTER treatment cell are cleared.
4. Confirm MASTER contains neither a `Φόρτος θεραπευτών` button nor `x/6` workload badges.
5. Then test a therapist absence that actually produces a replacement and confirm the success summary names patient, replacement therapist and time.
6. Only if all of these checks pass may G proceed to the first native `CloseOperationalDay` smoke.
