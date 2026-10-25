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
from rehab_excel.daily_replacement_facts import read_daily_replacement_facts  # noqa: E402
from rehab_excel.full_daily_projection import build_full_therapist_daily_plan  # noqa: E402
from rehab_excel.full_master_daily_projection import build_full_master_daily_plan  # noqa: E402
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402
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

    daily = read_daily_input(
        source,
        patients=patients,
        sessions=sessions,
        therapist_names=settings.therapist_names,
        ignore_rows=(source_row,),
        allow_empty=True,
    )
    replacements = read_daily_replacement_facts(
        source,
        sessions=sessions,
        target_date=target_date,
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
        "patch_count": len(plan.patches),
        "remaining_daily_input_changes": len(daily.source_changes),
        "active_replacements": len(replacements),
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
    print(f"Remaining daily changes: {result['remaining_daily_input_changes']}")
    print("Input unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
