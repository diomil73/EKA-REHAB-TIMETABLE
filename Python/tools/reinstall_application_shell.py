from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.application_shell_vba import install_application_shell  # noqa: E402
from rehab_excel.master_toolbar_vba import install_master_toolbar  # noqa: E402


class ApplicationShellRepairError(RuntimeError):
    pass


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def reinstall_application_shell(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> Path:
    """Create a fresh XLSM copy and reinstall the full app UI contract.

    The source remains immutable. The output receives the accepted MASTER toolbar,
    operational navigation buttons, ``modApplicationShell``, and workbook event
    hooks used to hide Excel chrome on open/activate.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise ApplicationShellRepairError(f"Source workbook not found: {source}")
    if source == output:
        raise ApplicationShellRepairError("Output must differ from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise ApplicationShellRepairError("Source and output must both be .xlsm")
    if not _has_vba(source):
        raise ApplicationShellRepairError("Source workbook has no VBA project")
    if output.exists() and not overwrite:
        raise ApplicationShellRepairError(f"Output already exists: {output}")

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)

    if sys.platform != "win32":
        try:
            output.unlink()
        except OSError:
            pass
        raise ApplicationShellRepairError("Application-shell repair requires Windows + Microsoft Excel")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        try:
            output.unlink()
        except OSError:
            pass
        raise ApplicationShellRepairError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        install_master_toolbar(workbook.VBProject, workbook)
        install_application_shell(workbook.VBProject, workbook=workbook)
        workbook.Save()
    except Exception as exc:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        workbook = None
        try:
            if output.exists():
                output.unlink()
        except OSError:
            pass
        raise ApplicationShellRepairError(f"Could not reinstall full application UI: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if not _has_vba(output):
        try:
            output.unlink()
        except OSError:
            pass
        raise ApplicationShellRepairError("Output lost its VBA project")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a fresh workbook copy with the full app UI contract reinstalled."
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        output = reinstall_application_shell(
            args.input,
            args.output,
            overwrite=args.overwrite,
        )
    except ApplicationShellRepairError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("APPLICATION UI REPAIR OK")
    print(f"Input: {args.input.resolve()}")
    print(f"Output: {output}")
    print("Source unchanged: True")
    print("VBA preserved: True")
    print("MASTER toolbar reinstalled: True")
    print("Operational navigation reinstalled: True")
    print("Application shell reinstalled: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
