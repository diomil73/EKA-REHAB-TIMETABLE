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
from rehab_excel.close_day_action_vba import CLOSE_DAY_MODULE_NAME  # noqa: E402
from rehab_excel.daily_change_ledger_vba import (  # noqa: E402
    MODULE_NAME as DAILY_CHANGE_LEDGER_MODULE_NAME,
    VBA_CODE as DAILY_CHANGE_LEDGER_BASE_CODE,
)
from rehab_excel.daily_input_change_ledger_vba import (  # noqa: E402
    MODULE_NAME as DAILY_INPUT_CHANGE_LEDGER_MODULE_NAME,
    install_daily_input_change_ledger,
)
from rehab_excel.daily_input_change_navigation_vba import (  # noqa: E402
    MODULE_NAME as DAILY_INPUT_CHANGE_NAV_MODULE_NAME,
    install_daily_input_change_navigation,
)
from rehab_excel.daily_replacement_ledger_sync_vba import (  # noqa: E402
    MODULE_NAME as REPLACEMENT_LEDGER_SYNC_MODULE_NAME,
    install_daily_replacement_ledger_sync,
)
from rehab_excel.daily_restore_in_session_vba import build_in_session_ledger_vba  # noqa: E402
from rehab_excel.double_booking_review_vba import (  # noqa: E402
    DOUBLE_BOOKING_FORM_NAME,
    DOUBLE_BOOKING_MODULE_NAME,
    install_double_booking_review,
)
from rehab_excel.global_replacement_action_vba import (  # noqa: E402
    GLOBAL_REPLACEMENT_MODULE_NAME,
    install_global_replacement_action,
)
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
from rehab_excel.replacement_operator_vba import (  # noqa: E402
    FORM_NAME as REPLACEMENT_OPERATOR_FORM_NAME,
    MODULE_NAME as REPLACEMENT_OPERATOR_MODULE_NAME,
    install_replacement_operator,
)
from rehab_excel.therapist_workload_strip_vba import (  # noqa: E402
    MODULE_NAME as WORKLOAD_MODULE_NAME,
    install_therapist_workload_strip,
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


def _component_present(vbproject, name: str) -> bool:
    try:
        _ = vbproject.VBComponents(name)
        return True
    except Exception:
        return False


def _ensure_double_booking_review(vbproject) -> None:
    module_present = _component_present(vbproject, DOUBLE_BOOKING_MODULE_NAME)
    form_present = _component_present(vbproject, DOUBLE_BOOKING_FORM_NAME)
    if module_present and form_present:
        return
    install_double_booking_review(vbproject)


def _prepare_clean_vba_copy(source: Path, temp_path: Path) -> None:
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

        components = (
            ("old menu module", MENU_MODULE_NAME),
            ("old menu form", MENU_FORM_NAME),
            ("old patient form", PATIENT_FORM_NAME),
            ("old therapist form", THERAPIST_FORM_NAME),
            ("old student form", STUDENT_FORM_NAME),
            ("old app shell module", APP_SHELL_MODULE_NAME),
            ("old MASTER toolbar module", MASTER_TOOLBAR_MODULE_NAME),
            ("old DAILY_INPUT action module", DAILY_INPUT_MODULE_NAME),
            ("old global replacement action module", GLOBAL_REPLACEMENT_MODULE_NAME),
            ("old replacement operator module", REPLACEMENT_OPERATOR_MODULE_NAME),
            ("old replacement operator form", REPLACEMENT_OPERATOR_FORM_NAME),
            ("old replacement ledger sync module", REPLACEMENT_LEDGER_SYNC_MODULE_NAME),
            ("old double-booking module", DOUBLE_BOOKING_MODULE_NAME),
            ("old double-booking form", DOUBLE_BOOKING_FORM_NAME),
            ("old workload module", WORKLOAD_MODULE_NAME),
            ("old daily change ledger module", DAILY_CHANGE_LEDGER_MODULE_NAME),
            ("old DAILY_INPUT change ledger module", DAILY_INPUT_CHANGE_LEDGER_MODULE_NAME),
            ("old DAILY_INPUT change navigation module", DAILY_INPUT_CHANGE_NAV_MODULE_NAME),
            ("old close-day action module", CLOSE_DAY_MODULE_NAME),
            ("old outpatient schedule module", OUTPATIENT_MODULE_NAME),
            ("old outpatient schedule form", OUTPATIENT_FORM_NAME),
        )
        for label, name in components:
            stage = f"removing {label}"
            _remove_component_if_present(vbproject, name)

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


def _install_operational_modules(workbook_path: Path) -> None:
    if sys.platform != "win32":
        raise RegistrationMenuVbaError("Operational module installation requires Windows + Excel")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RegistrationMenuVbaError("pywin32 is required") from exc

    excel = None
    workbook = None
    stage = "opening preview for operational modules"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        vbproject = workbook.VBProject

        stage = "installing global replacement action"
        install_global_replacement_action(vbproject)
        stage = "installing operator-controlled replacement workflow"
        install_replacement_operator(vbproject)
        stage = "verifying guided double-booking review"
        _ensure_double_booking_review(vbproject)
        stage = "installing therapist workload strip"
        install_therapist_workload_strip(vbproject, workbook)

        stage = "installing in-session daily change ledger"
        _remove_component_if_present(vbproject, DAILY_CHANGE_LEDGER_MODULE_NAME)
        ledger_module = vbproject.VBComponents.Add(1)
        ledger_module.Name = DAILY_CHANGE_LEDGER_MODULE_NAME
        ledger_module.CodeModule.AddFromString(
            build_in_session_ledger_vba(DAILY_CHANGE_LEDGER_BASE_CODE)
        )
        stage = "installing authoritative replacement ledger sync"
        install_daily_replacement_ledger_sync(vbproject)
        stage = "installing reversible DAILY_INPUT ledger"
        install_daily_input_change_ledger(vbproject)
        stage = "routing workload navigation through daily change sync"
        install_daily_input_change_navigation(vbproject, workbook)

        required = (
            GLOBAL_REPLACEMENT_MODULE_NAME,
            REPLACEMENT_OPERATOR_MODULE_NAME,
            REPLACEMENT_OPERATOR_FORM_NAME,
            REPLACEMENT_LEDGER_SYNC_MODULE_NAME,
            DOUBLE_BOOKING_MODULE_NAME,
            DOUBLE_BOOKING_FORM_NAME,
            WORKLOAD_MODULE_NAME,
            DAILY_CHANGE_LEDGER_MODULE_NAME,
            DAILY_INPUT_CHANGE_LEDGER_MODULE_NAME,
            DAILY_INPUT_CHANGE_NAV_MODULE_NAME,
        )
        missing = [name for name in required if not _component_present(vbproject, name)]
        if missing:
            raise RegistrationMenuVbaError(
                "Operational VBA components were not confirmed: " + ", ".join(missing)
            )
        workbook.Save()
    except RegistrationMenuVbaError:
        raise
    except Exception as exc:
        raise RegistrationMenuVbaError(
            f"Operational module installation failed during {stage}: {exc}"
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
        _install_operational_modules(output)
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
    print("Operational replacement/review/workload/ledger modules: present")
    print(f"Open Excel and run macro: {report.macro_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
