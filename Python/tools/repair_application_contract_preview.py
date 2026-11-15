from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import sys
import tempfile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.application_contract import (  # noqa: E402
    REQUIRED_APP_COMPONENTS,
    assert_application_contract,
)
from rehab_excel.outpatient_schedule_manager_vba import (  # noqa: E402
    FORM_NAME as OUTPATIENT_MANAGER_FORM,
    MODULE_NAME as OUTPATIENT_MANAGER_MODULE,
    install_outpatient_schedule_manager,
)


ALLOWED_MISSING = {OUTPATIENT_MANAGER_MODULE, OUTPATIENT_MANAGER_FORM}


def _component_names(vbproject) -> set[str]:
    return {
        str(vbproject.VBComponents.Item(index).Name)
        for index in range(1, int(vbproject.VBComponents.Count) + 1)
    }


def repair_preview(path: Path) -> tuple[tuple[str, ...], tuple[str, ...]]:
    if sys.platform != "win32":
        raise RuntimeError("Application-contract repair requires Windows + Microsoft Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("pywin32 is required") from exc

    path = path.resolve()
    if not path.exists() or path.suffix.casefold() != ".xlsm":
        raise RuntimeError("Target must be an existing .xlsm preview")

    backup_handle = tempfile.NamedTemporaryFile(
        prefix="eka_contract_repair_",
        suffix=".xlsm",
        dir=path.parent,
        delete=False,
    )
    backup_path = Path(backup_handle.name)
    backup_handle.close()
    shutil.copy2(path, backup_path)

    excel = None
    workbook = None
    before_missing: tuple[str, ...] = ()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject
        names = _component_names(vbproject)
        before_missing = tuple(name for name in REQUIRED_APP_COMPONENTS if name not in names)

        unexpected = tuple(name for name in before_missing if name not in ALLOWED_MISSING)
        if unexpected:
            raise RuntimeError(
                "Refusing narrow repair; unexpected required component(s) are missing: "
                + ", ".join(unexpected)
            )

        if before_missing:
            install_outpatient_schedule_manager(vbproject)

        names_after = _component_names(vbproject)
        after_missing = tuple(name for name in REQUIRED_APP_COMPONENTS if name not in names_after)
        if after_missing:
            raise RuntimeError(
                "Application contract still incomplete after repair: " + ", ".join(after_missing)
            )
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        excel.Quit()
        excel = None

        assert_application_contract(path)
        return before_missing, after_missing
    except Exception:
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
        shutil.copy2(backup_path, path)
        raise
    finally:
        backup_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Safely repair only the missing outpatient manager components in a disposable preview."
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    try:
        before_missing, after_missing = repair_preview(args.workbook)
    except Exception as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("APPLICATION CONTRACT REPAIR OK")
    print(f"Workbook: {args.workbook.resolve()}")
    print("Missing before: " + (", ".join(before_missing) if before_missing else "none"))
    print("Missing after: " + (", ".join(after_missing) if after_missing else "none"))
    print("Full application contract: verified")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
