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
- B, C, D, F, and G are frozen evidence artifacts and must not be repaired in place.
- E was never emitted after its fail-closed build failure.
- Close-day native testing remains blocked until the daily render/status and replacement-operator contracts below are restored on a fresh artifact.

## Native G findings

### Patient absence presentation is still wrong

The repeated patient-absence scenario confirmed that the current recovery path removes the patient's programme too aggressively from the operational presentation.

Agreed/render contract to restore:

- the original daily programme remains visible in its original slot,
- the absent/cancelled treatment is shown with a red strike-through/line rather than disappearing,
- the reason is displayed on the right in the agreed reason/status area,
- `MASTER_SCHEDULE`/daily status must not continue to describe the patient as `παρών` for that date,
- recurring/base programme remains untouched.

The repository `DAILY_RENDER_CONTRACT.md` already records the foundational rule that patient absence leaves the original patient line in place with strike-through. The latest native smoke proves that recovery work must reconnect that semantic render instead of clearing the visible programme.

### Therapist absence presentation is incomplete

Native therapist-absence flow did not expose the reason for the therapist absence. User-visible operational presentation must include the recorded reason and must not leak internal technical labels.

### Replacement operator UX is incomplete

The current automatic-proposal prompt exposes only Yes/No. Selecting `Όχι` produced no useful continuation. This is not the agreed operator workflow.

Durable replacement UX contract to restore/reconnect:

1. Replacement flow is triggered by therapist absence, not ordinary patient absence.
2. Per displaced patient, the operator has three explicit choices:
   - choose/confirm a replacement provider and time from valid available options,
   - postpone/suspend that patient's treatment for today,
   - defer the patient while continuing review/navigation rather than silently abandoning the flow.
3. Provider and time are explicit operator choices in manual mode; do not silently select the first option.
4. Each patient decision is explicitly confirmed and remains revisable with Previous/Next navigation.
5. A patient must not be considered completed without an explicit confirmed outcome.
6. There is also a global `Αυτόματη τακτοποίηση υπολοίπων` / best-solution action for all unresolved displaced patients together.
7. The automatic solver can be invoked at any point, including after some manual decisions. Confirmed manual decisions remain fixed/protected and the solver handles only the unresolved remainder.
8. Automatic resolution must reuse the established global solver and all locked rules, including hard feasibility constraints, scarcity protection, physiotherapy/reclining adjacency preference, workload balance, same-time preference, and weak 13:00 avoidance.
9. If automatic resolution cannot place everyone, fixed manual decisions remain intact and unresolved patients stay visible for operator review.
10. Automatic proposals require explicit operator confirmation before authoritative persistence.
11. Over-capacity/double-booking remains exceptional and requires explicit per-case review/confirmation.
12. After successful persistence, the operator must be shown who moved, to which therapist/provider, and at what time.

This contract is consistent with the existing global replacement decisions in `Documentation/DECISIONS.md`; the current native gap is primarily the thin operator UI/review reconnection, not a request to redesign the solver.

## Next recovery slice

Before any new native workbook:

1. restore patient-absence render semantics so the original line remains visible with red strike-through and reason, and date-specific status no longer reports `παρών`,
2. restore therapist-absence reason visibility,
3. reconnect the manual replacement review flow so `Όχι` from the automatic proposal enters the per-patient guided flow rather than doing nothing,
4. expose the three per-patient outcomes plus explicit Previous/Next/Confirm behavior,
5. keep global automatic resolution callable at any stage and make it respect already confirmed manual decisions,
6. preserve the existing solver/scoring rules rather than creating a second assignment engine,
7. cover the above with source regressions before creating the next uniquely named native artifact,
8. only after those native checks pass may `CloseOperationalDay` smoke resume.
