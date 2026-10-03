from __future__ import annotations

import argparse
from pathlib import Path
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]

PATIENT_FORM_NAME = "frmNewPatient"
BRIDGE_MARKERS = (
    "registration_bridge_cli.py",
    "BuildPatientRegistrationJson",
    "ResolveRegistrationBridgeScript",
)
OLD_MESSAGE_MARKER = "Τα στοιχεία είναι έτοιμα για αποστολή στο ασφαλές registration backend."


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Inspect the actual frmNewPatient VBA code stored in one .xlsm workbook."
    )
    parser.add_argument("--workbook", type=Path, required=True)
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    path = args.workbook.resolve()
    if not path.exists():
        print(f"WORKBOOK missing: {path}")
        return 2

    if sys.platform != "win32":
        print("Windows with Microsoft Excel is required.")
        return 2

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("pywin32 is required.")
        return 2

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)

        try:
            component = workbook.VBProject.VBComponents(PATIENT_FORM_NAME)
            module = component.CodeModule
            code = module.Lines(1, module.CountOfLines)
        except Exception as exc:
            print(f"INSPECTION STOP: {exc}")
            return 3

        print("PATIENT REGISTRATION BRIDGE INSPECTION")
        print(f"Workbook: {path}")
        print(f"Form: {PATIENT_FORM_NAME}")
        print(f"Code lines: {module.CountOfLines}")
        for marker in BRIDGE_MARKERS:
            print(f"Bridge marker [{marker}]: {marker in code}")
        print(f"Old MsgBox marker present: {OLD_MESSAGE_MARKER in code}")

        bridge_ok = all(marker in code for marker in BRIDGE_MARKERS)
        old_present = OLD_MESSAGE_MARKER in code
        print(f"BRIDGE WIRED IN STORED VBA: {bridge_ok and not old_present}")
        return 0 if bridge_ok and not old_present else 4
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


if __name__ == "__main__":
    raise SystemExit(main())
