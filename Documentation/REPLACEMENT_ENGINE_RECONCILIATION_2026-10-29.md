# Replacement engine reconciliation — 2026-10-29

## Why this checkpoint exists

A sequence of manual-operator V2 repair attempts produced repeated fail-closed errors such as `module declarations marker not found`. These errors came from runtime source patchers trying to transform already-modified VBA modules inside a disposable smoke workbook. They are not evidence that the replacement solver, feasibility rules, queue logic, or established manual/automatic replacement engine regressed.

## Stable boundary

Commit `560eace4bd7a0666ac838216f324da7080610987` is the current reconciliation boundary for the established replacement engine.

From that commit through branch head at the time of this checkpoint, the added work consisted only of new documentation, patch/repair helpers, and tests for the proposed Manual Operator V2 layer. The canonical replacement engine files were not modified by that V2 experiment.

The established engine therefore remains authoritative:

- `Python/rehab_excel/replacement_operator_vba.py`
- `Python/tools/prepare_replacement_operator_decision.py`
- the protected global replacement planner and its feasibility/capacity/policy rules
- authoritative `DAILY_REPLACEMENT_FACTS`
- the existing unresolved-session queue and sequential recalculation rules
- the existing automatic-all route
- the existing authoritative restore lifecycle

The lightweight manual-preparation fix remains valid: manual preparation must not invoke the full global solver merely to decorate a recommendation. Automatic mode remains the owner of the full global solve.

## Quarantined experiment

The following files are experimental V2 patching infrastructure and must not be used as the basis for further native acceptance until explicitly redesigned against a fresh canonical workbook build:

- `Python/rehab_excel/manual_replacement_operator_v2_patch.py`
- `Python/rehab_excel/manual_operator_v2_source_normalizer.py`
- `Python/tools/repair_manual_replacement_operator_v2.py`
- related V2 patch tests

The repeated `marker not found` failures are patcher/source-shape compatibility failures, not replacement-engine failures.

Do not continue adding marker-specific repairs to those scripts.

## Current native evidence

The established manual path reached the native candidate form promptly after the lightweight preparation fix. This proves the earlier CPU-heavy dependency on the full global solver was removed from manual preparation.

A separate native problem remains in the projection/rendering layer: after confirmation, Excel can raise a protected-sheet error while refreshing the visible change list on `THERAPIST_DAILY`. This must be treated as a renderer/protection problem only. It is not evidence to reopen replacement feasibility or solver logic.

## New operator UX requirements

The following are valid product requirements, but they must be implemented as a thin operator/presentation layer over the established engine rather than by rewriting the replacement engine:

1. `Αναβολή θεραπείας` must be available as an explicit per-day outcome for the current unresolved patient/session, using the existing date-specific treatment-suspension semantics.
2. Provider and time selection must be explicit. When one provider has multiple valid times, the operator chooses the exact time; the UI must not silently take the first available time.
3. Confirmation must not auto-advance. The operator explicitly chooses `Προηγούμενος ασθενής` / `Επόμενος ασθενής`.
4. A confirmed manual decision may be revisited and changed through authoritative undo/revalidation, with downstream availability recalculated.
5. `Αυτόματη τακτοποίηση υπολοίπων` must always be available and must run the existing global planner only on the unresolved remainder while respecting already confirmed manual outcomes.
6. The UI must expose clear progress/confirmed/busy state so the operator can distinguish working from stuck.

## Implementation rule from this point forward

No established engine file or rule is reopened without new evidence from a failing existing contract test or native proof tied to that layer.

Work proceeds in this order:

1. rebuild a fresh disposable smoke from the known canonical source path;
2. verify the canonical manual/automatic replacement tests before any UI work;
3. fix the `THERAPIST_DAILY` projection/protection issue as an isolated renderer change with its own regression test;
4. implement the new operator UX against canonical source, preferably by generating the complete authoritative form/module from source rather than mutating unknown live VBA with fragile string markers;
5. native smoke the new UI on a fresh disposable workbook;
6. only after native proof, fold the new UI into the normal preview/build pipeline.

## Explicit anti-regression rule

Do not use an already-repaired smoke workbook as the source of truth for code generation or patch assumptions. Disposable smoke workbooks are test artifacts, not canonical source.
