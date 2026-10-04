from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import sys
import tempfile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.application_shell_vba import APP_SHELL_MODULE_NAME  # noqa: E402
from rehab_excel.master_toolbar_vba import MASTER_TOOLBAR_MODULE_NAME  # noqa: E402
from rehab_excel.outpatient_schedule_form_vba import (  # noqa: E402
    FORM_NAME as OUTPATIENT_FORM_NAME,
    MODULE_NAME as OUTPATIENT_MODULE_NAME,
)
from rehab_excel.registration_menu_vba import (  # noqa: E402
    DEFAULT_PREVIEW_FILENAME,
    MENU_FORM_NAME,
    MENU_MODULE_NAME,
    DAILY_INPUT_MODULE_NAME,
    PATIENT_FORM_NAME,
    STUDENT_FORM_NAME,
    THERAPIST_FORM_NAME,
    RegistrationMenuVbaError,
    create_registration_menu_preview,
)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build a safe .xlsm preview containing the central registration UserForm."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / DEFAULT_PREVIEW_FILENAME,
    )
    parser.add_argument("--overwrite", action="store_true")
    return parser


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _prepare_clean_vba_copy(source: Path, temp_path: Path) -> None:
    """Remove replaceable forms/modules in one Excel session, then close it.

    Some Office builds keep a just-deleted UserForm name reserved until the
    workbook is saved and Excel closes. Rebuilding in a second session avoids
    the COM failure when assigning the same component name again.
    """

    if sys.platform != "win32":
        raise RegistrationMenuVbaError(
            "Registration menu installation requires Windows with Microsoft Excel"
        )

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RegistrationMenuVbaError(
            "pywin32 is required for registration menu installation"
        ) from exc

    shutil.copy2(source, temp_path)
    excel = None
    workbook = None
    stage = "opening temporary workbook"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(
            str(temp_path.resolve()), UpdateLinks=0, ReadOnly=False
        )
        stage = "accessing temporary VBA project"
        vbproject = workbook.VBProject
        _ = vbproject.VBComponents.Count

        stage = "removing old menu module"
        _remove_component_if_present(vbproject, MENU_MODULE_NAME)
        stage = "removing old menu form"
        _remove_component_if_present(vbproject, MENU_FORM_NAME)
        stage = "removing old patient form"
        _remove_component_if_present(vbproject, PATIENT_FORM_NAME)
        stage = "removing old therapist form"
        _remove_component_if_present(vbproject, THERAPIST_FORM_NAME)
        stage = "removing old student form"
        _remove_component_if_present(vbproject, STUDENT_FORM_NAME)
        stage = "removing old app shell module"
        _remove_component_if_present(vbproject, APP_SHELL_MODULE_NAME)
        stage = "removing old MASTER toolbar module"
        _remove_component_if_present(vbproject, MASTER_TOOLBAR_MODULE_NAME)
        stage = "removing old DAILY_INPUT action module"
        _remove_component_if_present(vbproject, DAILY_INPUT_MODULE_NAME)
        stage = "removing old outpatient schedule module"
        _remove_component_if_present(vbproject, OUTPATIENT_MODULE_NAME)
        stage = "removing old outpatient schedule form"
        _remove_component_if_present(vbproject, OUTPATIENT_FORM_NAME)

        stage = "saving temporary workbook"
        workbook.Save()
    except Exception as exc:
        raise RegistrationMenuVbaError(
            f"Excel registration menu pre-clean failed during {stage}: {exc}"
        ) from exc
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


def main() -> int:
    args = _build_parser().parse_args()
    source = args.source.resolve()
    output = args.output.resolve()

    if not source.exists():
        print(f"SAFETY STOP: Source workbook not found: {source}")
        return 2

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)

    temp_handle = tempfile.NamedTemporaryFile(
        prefix="eka_registration_menu_clean_",
        suffix=".xlsm",
        dir=output.parent,
        delete=False,
    )
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    temp_path.unlink(missing_ok=True)

    try:
        _prepare_clean_vba_copy(source, temp_path)
        report = create_registration_menu_preview(
            temp_path,
            output,
            overwrite=args.overwrite,
        )
        if _sha256(source) != source_before:
            output.unlink(missing_ok=True)
            raise RegistrationMenuVbaError("Original source workbook changed")
    except RegistrationMenuVbaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    finally:
        temp_path.unlink(missing_ok=True)

    print("REGISTRATION MENU PREVIEW OK")
    print(f"Preview: {report.output_path}")
    print("Source unchanged: True")
    print(f"VBA project present: {report.vba_present}")
    print(f"Menu module present: {report.module_present}")
    print(f"Menu form present: {report.form_present}")
    print(f"Open Excel and run macro: {report.macro_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
