from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.application_shell_vba import (  # noqa: E402
    APP_SHELL_MODULE_NAME,
    RETIRED_MASTER_SIDEBAR_FORM_NAME,
    install_application_shell,
)
from rehab_excel.patient_registration_scheduling_slice import (  # noqa: E402
    PatientRegistrationBridgeVbaError,
    create_patient_registration_bridge_preview,
)


def _refresh_application_shell(workbook_path: Path) -> None:
    """Install the current application shell into the final preview."""

    if sys.platform != "win32":
        raise PatientRegistrationBridgeVbaError(
            "Application-shell refresh requires Windows with Microsoft Excel"
        )

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientRegistrationBridgeVbaError("pywin32 is required") from exc

    excel = None
    workbook = None
    stage = "starting Excel"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        stage = "opening final preview"
        workbook = excel.Workbooks.Open(
            str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
        )

        stage = "installing current application shell"
        vbproject = workbook.VBProject
        _ = vbproject.VBComponents.Count
        install_application_shell(vbproject, workbook=workbook)

        stage = "verifying application shell"
        shell = vbproject.VBComponents(APP_SHELL_MODULE_NAME)
        shell_code = shell.CodeModule.Lines(1, shell.CodeModule.CountOfLines)
        if "Public Sub EnterApplicationShell()" not in str(shell_code):
            raise RuntimeError("EnterApplicationShell macro missing after shell refresh")
        if "Public Sub KeepApplicationShell()" not in str(shell_code):
            raise RuntimeError("KeepApplicationShell macro missing after shell refresh")

        try:
            vbproject.VBComponents(RETIRED_MASTER_SIDEBAR_FORM_NAME)
        except Exception:
            pass
        else:
            raise RuntimeError("Retired MASTER sidebar form still present after shell refresh")

        stage = "saving final preview"
        workbook.Save()
    except PatientRegistrationBridgeVbaError:
        raise
    except Exception as exc:
        raise PatientRegistrationBridgeVbaError(
            f"Final application-shell refresh failed during {stage}: {exc}"
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


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Create a safe .xlsm copy with patient registration followed by "
            "the recurring treatment scheduling form."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        report = create_patient_registration_bridge_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
        _refresh_application_shell(args.output)
    except PatientRegistrationBridgeVbaError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("PATIENT SCHEDULING SLICE PREVIEW OK")
    print(f"Source: {report.source_path}")
    print(f"Preview: {report.output_path}")
    print(f"Source unchanged: {report.source_unchanged}")
    print(f"VBA project present: {report.vba_present}")
    print(f"Patient form present: {report.form_present}")
    print(f"Registration bridge wired: {report.bridge_wired}")
    print(
        "NEXT: open only the preview, create a test patient, then continue into "
        "frmPatientScheduling and save one recurring treatment assignment."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
