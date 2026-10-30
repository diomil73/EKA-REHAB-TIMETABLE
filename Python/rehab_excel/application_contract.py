from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys


REQUIRED_APP_COMPONENTS: tuple[str, ...] = (
    "modApplicationShell",
    "modMasterToolbar",
    "modRegistrationMenu",
    "frmRegistrationMenu",
    "frmNewPatient",
    "frmNewTherapist",
    "frmNewStudent",
    "modPatientRegistrationBridge",
    "frmEditPatient",
    "modDailyInputAction",
    "modGlobalReplacementAction",
    "modTherapistWorkloadStrip",
    "modOutpatientSchedule",
    "frmOutpatientSchedule",
    "modOutpatientInSession",
    "modSchedulingConflict",
    "frmSchedulingConflict",
    "modOutpatientScheduleEdit",
    "modOutpatientScheduleManager",
    "frmOutpatientScheduleManager",
    "modCloseDayAction",
)


class ApplicationContractError(RuntimeError):
    pass


@dataclass(frozen=True)
class ApplicationContractReport:
    workbook_path: str
    component_names: tuple[str, ...]
    required_components: tuple[str, ...]


def assert_application_contract(workbook_path: str | Path) -> ApplicationContractReport:
    """Require the workbook to contain the established canonical application stack.

    This is intentionally stricter than merely checking for ``vbaProject.bin``.
    A workbook with a VBA project but missing one of the components produced by
    the canonical unified build pipeline is not a valid complete application
    artifact.
    """

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise ApplicationContractError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise ApplicationContractError("Application contract requires an .xlsm workbook")
    if sys.platform != "win32":
        raise ApplicationContractError(
            "Application contract verification requires Windows + Microsoft Excel"
        )

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ApplicationContractError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)

        if not bool(workbook.HasVBProject):
            raise ApplicationContractError("Workbook has no VBA project")

        try:
            vbproject = workbook.VBProject
            component_names = tuple(
                sorted(
                    str(vbproject.VBComponents.Item(index).Name)
                    for index in range(1, int(vbproject.VBComponents.Count) + 1)
                )
            )
        except Exception as exc:
            raise ApplicationContractError(
                "Could not inspect VBA components. Enable Trust access to the VBA project object model."
            ) from exc

        missing = tuple(
            name for name in REQUIRED_APP_COMPONENTS if name not in set(component_names)
        )
        if missing:
            raise ApplicationContractError(
                "Workbook is not a complete application artifact; missing VBA component(s): "
                + ", ".join(missing)
            )

        return ApplicationContractReport(
            workbook_path=str(path),
            component_names=component_names,
            required_components=REQUIRED_APP_COMPONENTS,
        )
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
