from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_reader import read_daily_input  # noqa: E402
from rehab_excel.daily_preview_composer import compose_outpatient_daily_plan  # noqa: E402
from rehab_excel.daily_replacement_facts import (  # noqa: E402
    read_daily_replacement_fact_rows,
    read_daily_replacement_facts,
)
from rehab_excel.full_daily_projection import build_full_therapist_daily_plan  # noqa: E402
from rehab_excel.full_master_daily_projection import build_full_master_daily_plan  # noqa: E402
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402
from rehab_excel.therapist_absence_restore import (  # noqa: E402
    dependent_active_replacement_ids,
    dependent_generated_suspension_ids,
    released_therapist_session_ids,
)
from rehab_excel.writeback import build_write_plan, validate_write_plan  # noqa: E402
from tools.prepare_daily_input_in_place import _patch_line, _target_date  # noqa: E402


class DailyInputRestorePreparationError(RuntimeError):
    pass


def _write_response(path: Path | None, payload: dict[str, object]) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8-sig",
    )


def _source_rows_from_cancellation_ids(values: tuple[str, ...]) -> tuple[int, ...]:
    rows: list[int] = []
    for value in values:
        prefix = "daily-input:"
        if not value.startswith(prefix):
            raise DailyInputRestorePreparationError(
                f"Generated suspension has unsupported source identity: {value!r}"
            )
        raw_row = value[len(prefix) :]
        if not raw_row.isdigit():
            raise DailyInputRestorePreparationError(
                f"Generated suspension has invalid source row identity: {value!r}"
            )
        rows.append(int(raw_row))
    return tuple(sorted(set(rows)))


def prepare_daily_input_restore_payload(
    workbook_path: str | Path,
    payload_path: str | Path,
    *,
    source_row: int,
) -> dict[str, object]:
    source = Path(workbook_path).resolve()
    payload = Path(payload_path).resolve()
    if not source.exists():
        raise DailyInputRestorePreparationError(f"Workbook not found: {source}")
    if source_row < 1:
        raise DailyInputRestorePreparationError("source_row must be positive")

    source_hash = file_sha256(source)
    assert_application_contract(source)
    target_date = _target_date(source)
    patients = read_patient_registry(source)
    settings = read_settings(source)
    base_entries = read_unified_base_schedule(source, patients=patients)
    sessions = materialize_sessions_for_date(base_entries, target_date)

    before = read_daily_input(
        source,
        patients=patients,
        sessions=sessions,
        therapist_names=settings.therapist_names,
        allow_empty=True,
    )
    selected_changes = tuple(
        change for change in before.source_changes if change.source_row == source_row
    )
    if len(selected_changes) != 1:
        raise DailyInputRestorePreparationError(
            f"DAILY_INPUT row {source_row} is not one active source change"
        )
    selected_change = selected_changes[0]

    after_selected = read_daily_input(
        source,
        patients=patients,
        sessions=sessions,
        therapist_names=settings.therapist_names,
        ignore_rows=(source_row,),
        allow_empty=True,
    )

    replacement_restore_ids: tuple[str, ...] = ()
    generated_suspension_rows: tuple[int, ...] = ()
    released_session_ids: tuple[str, ...] = ()

    if selected_change.change_type == "therapist_absence":
        released_session_ids = released_therapist_session_ids(
            sessions=sessions,
            before_absences=before.absences,
            after_absences=after_selected.absences,
        )
        replacement_rows = read_daily_replacement_fact_rows(
            source,
            target_date=target_date,
            include_restored=True,
        )
        replacement_restore_ids = dependent_active_replacement_ids(
            replacement_rows,
            released_session_ids=released_session_ids,
        )
        generated_suspension_ids = dependent_generated_suspension_ids(
            before.cancellations,
            released_session_ids=released_session_ids,
        )
        generated_suspension_rows = _source_rows_from_cancellation_ids(
            generated_suspension_ids
        )

    ignored_rows = tuple(sorted({source_row, *generated_suspension_rows}))
    daily = read_daily_input(
        source,
        patients=patients,
        sessions=sessions,
        therapist_names=settings.therapist_names,
        ignore_rows=ignored_rows,
        allow_empty=True,
    )
    replacements = tuple(
        replacement
        for replacement in read_daily_replacement_facts(
            source,
            sessions=sessions,
            target_date=target_date,
        )
        if replacement.replacement_id not in set(replacement_restore_ids)
    )
    states = build_daily_session_states(
        sessions,
        absences=daily.absences,
        replacements=replacements,
        cancellations=daily.cancellations,
        target_date=target_date,
    )

    therapist_plan = build_full_therapist_daily_plan(
        source,
        base_entries=base_entries,
        states=states,
        sessions=sessions,
        patients=patients,
    )
    therapist_plan = compose_outpatient_daily_plan(
        therapist_plan,
        states=states,
        patients=patients,
    )
    master_plan = build_full_master_daily_plan(
        source,
        base_entries=base_entries,
        states=states,
        patients=patients,
    )

    plan = build_write_plan(
        source,
        (*therapist_plan.patches, *master_plan.patches),
    )
    validate_write_plan(plan)

    payload.parent.mkdir(parents=True, exist_ok=True)
    payload.write_text(
        "\n".join(_patch_line(patch) for patch in plan.patches) + "\n",
        encoding="utf-8-sig",
    )

    if file_sha256(source) != source_hash:
        payload.unlink(missing_ok=True)
        raise DailyInputRestorePreparationError(
            "Workbook changed during DAILY_INPUT restore validation"
        )

    return {
        "ok": True,
        "date": target_date.isoformat(),
        "source_row": source_row,
        "source_change_type": selected_change.change_type,
        "patch_count": len(plan.patches),
        "remaining_daily_input_changes": len(daily.source_changes),
        "active_replacements": len(replacements),
        "released_session_count": len(released_session_ids),
        "replacement_restore_count": len(replacement_restore_ids),
        "replacement_restore_ids": "|".join(replacement_restore_ids),
        "generated_suspension_count": len(generated_suspension_rows),
        "generated_suspension_rows": ",".join(
            str(row) for row in generated_suspension_rows
        ),
        "source_unchanged": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Prepare a full in-place rebuild after restoring one DAILY_INPUT change."
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument("--source-row", required=True, type=int)
    parser.add_argument("--response", type=Path)
    args = parser.parse_args()

    try:
        result = prepare_daily_input_restore_payload(
            args.input,
            args.payload,
            source_row=args.source_row,
        )
    except Exception as exc:
        args.payload.unlink(missing_ok=True)
        _write_response(args.response, {"ok": False, "error": str(exc)})
        print(f"SAFETY STOP: {exc}")
        return 2

    _write_response(args.response, result)
    print("DAILY INPUT RESTORE PAYLOAD VERIFIED")
    print(f"Date: {result['date']}")
    print(f"Source row: {result['source_row']}")
    print(f"Patches: {result['patch_count']}")
    print(f"Released sessions: {result['released_session_count']}")
    print(f"Replacement facts to restore: {result['replacement_restore_count']}")
    print(f"Generated suspensions to clear: {result['generated_suspension_count']}")
    print(f"Remaining daily changes: {result['remaining_daily_input_changes']}")
    print("Input unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
