# Daily change restore recovery checkpoint — 2026-10-31

## Native H evidence

Native H is frozen as evidence. Do not continue therapist-absence replacement testing or `CloseOperationalDay` on H.

Patient-absence smoke was operationally correct except for two UX/reversibility gaps reported by the operator:

1. The MASTER status/reason text was also visually struck through. The reason must remain immediately readable.
2. The daily change needs a supported restore workflow from the daily schedule, with one-row, selected-row and all-row restore actions.

## MASTER reason presentation fix

The patient-status overlay must be red/readable but not struck through. The daily non-active treatment cell may remain struck through.

Source fix:

- `430c2e90f65ee97cfd28577dd47af0aec56fd58f` — explicitly sets `master_daily_patient_status_overlay` to `Strikethrough = False` and verifies it by native semantic readback.
- `562530c0f0b9c579de8ee64907bd505204f749ca` — regression coverage for the readable reason contract.

## Recovered historical ledger UX contract

The repository already contains the previously designed daily change ledger UX in `Python/rehab_excel/daily_change_ledger_vba.py`.

The useful UX contract is retained:

- change list rendered below `THERAPIST_DAILY`;
- columns for patient, from, to, reason, state and action;
- per-row `Επαναφορά`;
- checkboxes for active rows;
- `Επαναφορά επιλεγμένων`;
- `Επαναφορά όλων`;
- restored rows remain visible as grey/italic history with state `Επαναφέρθηκε`.

The old restore transport is **not** authoritative for the recovery build. Even its later in-session wrapper still created a separate restore preview and synchronized operational sheets back into the open workbook. The current recovery architecture is stricter:

`validate -> prepare read-only payload -> in-place native apply -> semantic readback -> rollback on failure -> save same workbook`

Therefore we recover the ledger UX but replace the old preview/sync restore backend.

## Generalized change model

Restore is broader than replacement facts. The operator requirement is that everyone affected by a daily operational change appears in the change ledger.

`daily_input_reader.py` already exposes reversible source facts through `DailyInputSourceChange`:

- stable `change_id`;
- `source_row`;
- `change_type`;
- `subject_id`;
- display name;
- reason;
- all-day/single-time metadata.

Current source change types include:

- `patient_absence`;
- `session_cancellation`;
- `therapist_absence`.

A restore must reverse the authoritative source fact, not copy cells from an old snapshot.

## First in-place restore preparer

Commit:

- `8573572df78e1d3cff6d1847f04a3c3e9ffdedca` — `Python/tools/prepare_daily_change_restore_in_place.py`.

Current deliberately bounded scope:

- supports one active `patient_absence` or `session_cancellation` change;
- resolves exactly one `change_id` from current `DAILY_INPUT` source changes;
- re-reads daily input while ignoring only that exact `source_row`;
- preserves every other daily input row;
- preserves currently ACTIVE replacement facts;
- rebuilds daily session truth;
- rebuilds patient-centric/OUTPATIENT projection;
- rebuilds MASTER daily overlay;
- validates the resulting write plan;
- emits an in-place payload containing:
  - `SOURCE_CLEAR` for the exact DAILY_INPUT patient row `A:E`;
  - `RESTORE_CHANGE` for the immutable change id;
  - validated `PATCH` rows for operational projections;
- verifies source SHA is unchanged;
- performs no workbook writes, no workbook copy, no Excel COM mutation and no close/reopen.

`therapist_absence` restore intentionally fails closed at preparation time. It requires coordinated rollback of replacement facts generated for that absence. Those facts must never be left ACTIVE after the therapist source absence is removed.

Regression commit:

- `0b85901df027ad087423107352f9fc14ac135e49` — focused source/architecture tests for the new preparer.

## Next gate

Before adding VBA apply/UI integration, run:

```powershell
git pull --ff-only origin work/history-statistics-rebuild

python -m pytest `
  Tests/test_prepare_daily_change_restore_in_place.py `
  Tests/test_operator_visibility_patch_vba.py `
  Tests/test_master_daily_overlay.py `
  Tests/test_prepare_daily_input_in_place.py `
  Tests/test_daily_input_action_vba.py `
  Tests/test_production_actions_stay_in_session.py -q
```

If green, the next bounded slice is:

1. add native `SOURCE_CLEAR` + `RESTORE_CHANGE` apply/readback/rollback handling;
2. render the generalized change ledger below `THERAPIST_DAILY`;
3. enable per-row restore first for patient absence/session cancellation;
4. add selected/all restore only when all selected rows are compatible with one atomic request;
5. keep therapist-absence restore disabled until related replacement-fact rollback is authoritative;
6. install the recovered ledger UI through the canonical unified builder;
7. build a fresh I artifact only after the focused source gate is green.

No production workbook is modified by this checkpoint or preparer.
