from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_core.daily_statistics import calculate_daily_treatment_statistics  # noqa: E402
from rehab_excel.daily_input_reader import DailyInputReadError, read_daily_input  # noqa: E402
from rehab_excel.daily_preview_composer import (  # noqa: E402
    DailyPreviewCompositionError,
    compose_outpatient_daily_plan,
)
from rehab_excel.daily_replacement_fact_batch import (  # noqa: E402
    restore_daily_replacement_facts,
)
from rehab_excel.daily_replacement_facts import (  # noqa: E402
    DailyReplacementFactError,
    read_daily_replacement_fact_rows,
    read_daily_replacement_facts,
)
from rehab_excel.daily_statistics_writer import (  # noqa: E402
    DailyStatisticsWriteError,
    write_daily_statistics,
)
from rehab_excel.master_daily_overlay import (  # noqa: E402
    MasterDailyOverlayError,
    build_master_daily_overlay_plan,
    compose_master_daily_overlay,
)
from rehab_excel.master_projection_writer import (  # noqa: E402
    MasterProjectionWriteError,
    refresh_master_projection_in_place,
)
from rehab_excel.native_excel import apply_write_plan_to_copy  # noqa: E402
from rehab_excel.outpatient_presentation import OutpatientPresentationError  # noqa: E402
from rehab_excel.outpatient_schedule_source import (  # noqa: E402
    OutpatientScheduleSourceError,
    read_unified_base_schedule,
)
from rehab_excel.patient_centric_preview import (  # noqa: E402
    PatientCentricPreviewError,
    build_patient_centric_preview_plan,
)
from rehab_excel.patient_registry_source import (  # noqa: E402
    PatientRegistrySourceError,
    read_patient_registry,
)
from tools.restore_daily_replacement_fact import (  # noqa: E402
    AuthoritativeDailyRestoreError,
    _has_vba,
    _mark_ledger_restored,
    _read_daily_input_date,
    _sha256,
)


def _write_response(path: Path | None, payload: dict[str, object]) -> None:
    """Write the small UTF-8 JSON bridge response when requested by Excel/VBA."""

    if path is None:
        return
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False),
        encoding="utf-8-sig",
    )


def _requested_restore_ids(
    workbook_path: Path,
    *,
    target_date,
    explicit_ids: tuple[str, ...],
    restore_all: bool,
) -> tuple[str, ...]:
    if restore_all:
        facts = read_daily_replacement_fact_rows(
            workbook_path,
            target_date=target_date,
            include_restored=False,
        )
        ids = tuple(fact.replacement_id for fact in facts)
        if not ids:
            raise AuthoritativeDailyRestoreError("No ACTIVE replacement facts to restore")
        return ids
    cleaned = tuple(value.strip() for value in explicit_ids if value.strip())
    if not cleaned:
        raise AuthoritativeDailyRestoreError("At least one --replacement-id is required")
    return cleaned


def restore_replacements_and_recompute(
    source_path: str | Path,
    output_path: str | Path,
    *,
    replacement_ids: tuple[str, ...] = (),
    restore_all: bool = False,
    overwrite: bool = False,
) -> tuple[int, int, int, bool]:
    """Restore a validated fact set, then rebuild effective daily truth once."""

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise AuthoritativeDailyRestoreError(f"Source workbook not found: {source}")
    if source == output:
        raise AuthoritativeDailyRestoreError("Output must differ from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise AuthoritativeDailyRestoreError("Source and output must both be .xlsm")
    if restore_all and replacement_ids:
        raise AuthoritativeDailyRestoreError(
            "Use either explicit replacement ids or restore_all, not both"
        )
    if output.exists() and not overwrite:
        raise AuthoritativeDailyRestoreError(f"Output already exists: {output}")

    source_before = _sha256(source)
    stage = output.with_name(f".{output.stem}.bulk-restore-stage{output.suffix}")
    if stage.exists():
        stage.unlink()
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, stage)

    try:
        target_date = _read_daily_input_date(stage)
        wanted_ids = _requested_restore_ids(
            stage,
            target_date=target_date,
            explicit_ids=replacement_ids,
            restore_all=restore_all,
        )
        restored_ids = restore_daily_replacement_facts(stage, wanted_ids)

        patients = read_patient_registry(stage)
        base_entries = read_unified_base_schedule(stage)
        sessions = materialize_sessions_for_date(base_entries, target_date)
        daily = read_daily_input(stage, patients=patients, sessions=sessions)
        replacements = read_daily_replacement_facts(
            stage,
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
        statistics = calculate_daily_treatment_statistics(states)

        refresh_master_projection_in_place(stage)
        preview = build_patient_centric_preview_plan(
            stage,
            base_entries=base_entries,
            states=states,
            sessions=sessions,
            patients=patients,
            rebuild_multi_member_groups=True,
        )
        composed_plan = compose_outpatient_daily_plan(
            preview.write_plan,
            states=states,
            patients=patients,
        )
        master_overlay = build_master_daily_overlay_plan(
            stage,
            base_entries=base_entries,
            states=states,
            patients=patients,
        )
        composed_plan = compose_master_daily_overlay(composed_plan, master_overlay)
        report = apply_write_plan_to_copy(
            composed_plan,
            output,
            overwrite=overwrite,
        )
        write_daily_statistics(output, target_date, statistics)
        for replacement_id in restored_ids:
            _mark_ledger_restored(output, replacement_id)

        source_after = _sha256(source)
        if source_before != source_after:
            raise AuthoritativeDailyRestoreError(
                "Source workbook changed during bulk restore preview"
            )
        vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)
        if not report.source_unchanged or not vba_preserved:
            raise AuthoritativeDailyRestoreError(
                "Bulk restore preview failed source/VBA preservation checks"
            )
        return (
            len(restored_ids),
            len(replacements),
            statistics.unresolved_sessions,
            statistics.is_finalized,
        )
    finally:
        try:
            stage.unlink()
        except OSError:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Restore selected or all ACTIVE daily replacement facts atomically, "
            "then recompute the saved day once."
        )
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--replacement-id", action="append", default=[])
    parser.add_argument("--all-active", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument(
        "--response",
        type=Path,
        help="Optional UTF-8 JSON response file for the Excel/VBA authoritative bridge.",
    )
    args = parser.parse_args()

    input_path = args.input.resolve()
    source_sha256_before = _sha256(input_path) if input_path.exists() else ""

    try:
        restored, active, unresolved, finalized = restore_replacements_and_recompute(
            args.input,
            args.output,
            replacement_ids=tuple(args.replacement_id),
            restore_all=args.all_active,
            overwrite=args.overwrite,
        )
    except (
        AuthoritativeDailyRestoreError,
        DailyInputReadError,
        DailyReplacementFactError,
        DailyStatisticsWriteError,
        PatientCentricPreviewError,
        PatientRegistrySourceError,
        OutpatientScheduleSourceError,
        OutpatientPresentationError,
        DailyPreviewCompositionError,
        MasterDailyOverlayError,
        MasterProjectionWriteError,
        ValueError,
        KeyError,
    ) as exc:
        _write_response(
            args.response,
            {
                "ok": False,
                "error": str(exc),
                "source_sha256_before": source_sha256_before,
            },
        )
        print(f"SAFETY STOP: {exc}")
        return 2

    output_path = args.output.resolve()
    _write_response(
        args.response,
        {
            "ok": True,
            "output_path": str(output_path),
            "source_sha256_before": source_sha256_before,
            "restored_replacement_facts": restored,
            "active_replacements_after_restore": active,
            "unresolved_sessions_after_recompute": unresolved,
            "finalized_after_recompute": finalized,
            "source_unchanged": True,
            "vba_preserved": True,
        },
    )

    print("AUTHORITATIVE BULK DAILY RESTORE OK")
    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Restored replacement facts: {restored}")
    print(f"Active replacements after restore: {active}")
    print(f"Unresolved sessions after recompute: {unresolved}")
    print(f"Finalized after recompute: {'ΝΑΙ' if finalized else 'ΟΧΙ'}")
    print("Source unchanged: True")
    print("VBA preserved: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
