from __future__ import annotations

import argparse
import sys
from pathlib import Path
from time import perf_counter

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

import tools.restore_daily_replacement_facts as restore_tool  # noqa: E402


def _seconds(started_at: float) -> float:
    return perf_counter() - started_at


def _profiled_apply_factory(timings: dict[str, float]):
    def profiled_apply(self, workbook_path: Path, patches) -> None:
        if sys.platform != "win32":
            raise restore_tool.AuthoritativeDailyRestoreError(
                "Restore output write-back requires Windows + Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise restore_tool.AuthoritativeDailyRestoreError("pywin32 is required") from exc

        excel = None
        workbook = None
        phase_started = perf_counter()
        try:
            started = perf_counter()
            excel = win32com.client.DispatchEx("Excel.Application")
            timings["excel_dispatch"] = _seconds(started)

            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False

            started = perf_counter()
            workbook = excel.Workbooks.Open(
                str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
            )
            timings["workbook_open"] = _seconds(started)

            started = perf_counter()
            self._apply_master_projection(workbook)
            timings["master_projection"] = _seconds(started)

            started = perf_counter()
            patch_count = 0
            for patch in patches:
                self._apply_patch(workbook, patch)
                patch_count += 1
            timings["daily_patches"] = _seconds(started)
            timings["daily_patch_count"] = float(patch_count)

            started = perf_counter()
            self._write_statistics(workbook)
            timings["statistics"] = _seconds(started)

            started = perf_counter()
            self._mark_replacement_fact_rows(workbook)
            timings["replacement_fact_rows"] = _seconds(started)

            started = perf_counter()
            self._mark_ledger_rows(workbook)
            timings["ledger_rows"] = _seconds(started)

            started = perf_counter()
            try:
                excel.Run(f"'{workbook.Name}'!RefreshDailyChangeLedgerView")
            except Exception:
                pass
            timings["ledger_refresh"] = _seconds(started)

            started = perf_counter()
            workbook.Save()
            timings["workbook_save"] = _seconds(started)
        except restore_tool.AuthoritativeDailyRestoreError:
            raise
        except Exception as exc:
            raise restore_tool.AuthoritativeDailyRestoreError(
                f"Could not write authoritative bulk restore output: {exc}"
            ) from exc
        finally:
            started = perf_counter()
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            timings["workbook_close"] = _seconds(started)

            started = perf_counter()
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            timings["excel_quit"] = _seconds(started)
            timings["excel_output_total"] = _seconds(phase_started)

    return profiled_apply


def _timed_wrapper(
    timings: dict[str, float],
    key: str,
    func,
):
    def wrapped(*args, **kwargs):
        started = perf_counter()
        try:
            return func(*args, **kwargs)
        finally:
            timings[key] = timings.get(key, 0.0) + _seconds(started)

    return wrapped


def profile_restore_backend(
    input_path: Path,
    output_path: Path,
    *,
    replacement_ids: tuple[str, ...] = (),
    restore_all: bool = False,
    overwrite: bool = False,
) -> tuple[tuple[int, int, int, bool], dict[str, float]]:
    """Run the real authoritative restore while timing backend and Excel phases."""

    timings: dict[str, float] = {}
    originals: list[tuple[object, str, object]] = []

    def patch(owner, name: str, key: str) -> None:
        original = getattr(owner, name)
        originals.append((owner, name, original))
        setattr(owner, name, _timed_wrapper(timings, key, original))

    original_apply = restore_tool._RestoreOutputBackend.apply
    restore_tool._RestoreOutputBackend.apply = _profiled_apply_factory(timings)

    patch(restore_tool.shutil, "copy2", "stage_copy")
    patch(restore_tool, "_sha256", "sha256_checks")
    patch(restore_tool, "_read_daily_input_date", "read_target_date")
    patch(restore_tool, "read_daily_replacement_fact_rows", "read_replacement_fact_rows")
    patch(restore_tool, "_requested_restore_ids", "resolve_restore_ids")
    patch(restore_tool, "validate_batch_restore_request", "validate_restore_ids")
    patch(restore_tool, "read_patient_registry", "read_patients")
    patch(restore_tool, "read_unified_base_schedule", "read_base_schedule")
    patch(restore_tool, "materialize_sessions_for_date", "materialize_sessions")
    patch(restore_tool, "read_daily_input", "read_daily_input")
    patch(restore_tool, "_active_replacements_after_restore", "derive_replacement_facts")
    patch(restore_tool, "build_daily_session_states", "build_daily_states")
    patch(restore_tool, "calculate_daily_treatment_statistics", "calculate_statistics")
    patch(restore_tool, "build_master_projection", "build_master_projection")
    patch(restore_tool, "build_patient_centric_preview_plan", "build_patient_preview")
    patch(restore_tool, "compose_outpatient_daily_plan", "compose_outpatient_plan")
    patch(restore_tool, "build_master_daily_overlay_plan", "build_master_overlay")
    patch(restore_tool, "compose_master_daily_overlay", "compose_master_overlay")
    patch(restore_tool, "apply_write_plan_to_copy", "apply_write_plan_total")
    patch(restore_tool, "_has_vba", "vba_checks")

    total_started = perf_counter()
    try:
        result = restore_tool.restore_replacements_and_recompute(
            input_path,
            output_path,
            replacement_ids=replacement_ids,
            restore_all=restore_all,
            overwrite=overwrite,
        )
    finally:
        restore_tool._RestoreOutputBackend.apply = original_apply
        for owner, name, original in reversed(originals):
            setattr(owner, name, original)

    timings["backend_total"] = _seconds(total_started)
    timings["pre_excel_and_verification"] = max(
        0.0,
        timings["backend_total"] - timings.get("excel_output_total", 0.0),
    )
    timed_pre_excel = sum(
        timings.get(key, 0.0)
        for key in (
            "stage_copy",
            "sha256_checks",
            "read_target_date",
            "read_replacement_fact_rows",
            "resolve_restore_ids",
            "validate_restore_ids",
            "read_patients",
            "read_base_schedule",
            "materialize_sessions",
            "read_daily_input",
            "derive_replacement_facts",
            "build_daily_states",
            "calculate_statistics",
            "build_master_projection",
            "build_patient_preview",
            "compose_outpatient_plan",
            "build_master_overlay",
            "compose_master_overlay",
            "vba_checks",
        )
    )
    timings["pre_excel_unattributed"] = max(
        0.0,
        timings["pre_excel_and_verification"] - timed_pre_excel,
    )
    return result, timings


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Profile authoritative daily restore backend phases without changing "
            "the production restore contract."
        )
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--replacement-id", action="append", default=[])
    parser.add_argument("--all-active", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    result, timings = profile_restore_backend(
        args.input.resolve(),
        args.output.resolve(),
        replacement_ids=tuple(args.replacement_id),
        restore_all=args.all_active,
        overwrite=args.overwrite,
    )
    restored, active, unresolved, finalized = result

    print("AUTHORITATIVE RESTORE BACKEND PROFILE")
    print(f"Restored facts: {restored}")
    print(f"Active replacements after: {active}")
    print(f"Unresolved after: {unresolved}")
    print(f"Finalized after: {'ΝΑΙ' if finalized else 'ΟΧΙ'}")
    print()
    print("Phase timings (seconds):")
    for key in (
        "backend_total",
        "pre_excel_and_verification",
        "stage_copy",
        "sha256_checks",
        "read_target_date",
        "read_replacement_fact_rows",
        "resolve_restore_ids",
        "validate_restore_ids",
        "read_patients",
        "read_base_schedule",
        "materialize_sessions",
        "read_daily_input",
        "derive_replacement_facts",
        "build_daily_states",
        "calculate_statistics",
        "build_master_projection",
        "build_patient_preview",
        "compose_outpatient_plan",
        "build_master_overlay",
        "compose_master_overlay",
        "vba_checks",
        "pre_excel_unattributed",
        "apply_write_plan_total",
        "excel_output_total",
        "excel_dispatch",
        "workbook_open",
        "master_projection",
        "daily_patches",
        "statistics",
        "replacement_fact_rows",
        "ledger_rows",
        "ledger_refresh",
        "workbook_save",
        "workbook_close",
        "excel_quit",
    ):
        if key in timings:
            print(f"{key}: {timings[key]:.3f}")
    if "daily_patch_count" in timings:
        print(f"daily_patch_count: {int(timings['daily_patch_count'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
