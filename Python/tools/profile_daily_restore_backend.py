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


def profile_restore_backend(
    input_path: Path,
    output_path: Path,
    *,
    replacement_ids: tuple[str, ...] = (),
    restore_all: bool = False,
    overwrite: bool = False,
) -> tuple[tuple[int, int, int, bool], dict[str, float]]:
    """Run the real authoritative restore while timing the native Excel output phases."""

    timings: dict[str, float] = {}
    original_apply = restore_tool._RestoreOutputBackend.apply
    restore_tool._RestoreOutputBackend.apply = _profiled_apply_factory(timings)
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
    timings["backend_total"] = _seconds(total_started)
    timings["pre_excel_and_verification"] = max(
        0.0,
        timings["backend_total"] - timings.get("excel_output_total", 0.0),
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
        "excel_output_total",
        "excel_dispatch",
        "workbook_open",
        "master_projection",
        "daily_patches",
        "statistics",
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
