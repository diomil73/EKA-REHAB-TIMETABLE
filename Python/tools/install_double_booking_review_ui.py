from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.double_booking_daily_projection_vba import (  # noqa: E402
    DOUBLE_BOOKING_PROJECTION_MODULE_NAME,
    install_double_booking_daily_projection,
)
from rehab_excel.double_booking_review_vba import (  # noqa: E402
    DOUBLE_BOOKING_FORM_NAME,
    DOUBLE_BOOKING_MODULE_NAME,
    FORM_CODE,
    install_double_booking_review,
)


class DoubleBookingUiInstallError(RuntimeError):
    pass


def _component_present(vbproject, name: str) -> bool:
    try:
        _ = vbproject.VBComponents(name)
        return True
    except Exception:
        return False


def _projection_aware_form_code() -> str:
    original = '''    If modDoubleBookingReview.ApplySelectedDoubleBooking( _
        lstOptions.List(idx, 0), _
        lstOptions.List(idx, 3), _
        lstOptions.List(idx, 5)) Then
        Unload Me
    End If'''
    replacement = '''    If modDoubleBookingReview.ApplySelectedDoubleBooking( _
        lstOptions.List(idx, 0), _
        lstOptions.List(idx, 3), _
        lstOptions.List(idx, 5)) Then
        If modDoubleBookingDailyProjection.ProjectDoubleBookingToTherapistDaily( _
            lstOptions.List(idx, 1), _
            lstOptions.List(idx, 4), _
            lstOptions.List(idx, 5), _
            lstOptions.List(idx, 0)) Then
            Unload Me
        End If
    End If'''
    if original not in FORM_CODE:
        raise DoubleBookingUiInstallError(
            "Double-booking form confirmation block no longer matches installer contract"
        )
    return FORM_CODE.replace(original, replacement, 1)


def _replace_form_code(vbproject) -> None:
    form = vbproject.VBComponents(DOUBLE_BOOKING_FORM_NAME)
    code_module = form.CodeModule
    if int(code_module.CountOfLines) > 0:
        code_module.DeleteLines(1, int(code_module.CountOfLines))
    code_module.AddFromString(_projection_aware_form_code())


def install_ui(workbook_path: Path) -> tuple[int, bool, bool, bool]:
    if sys.platform != "win32":
        raise DoubleBookingUiInstallError("Windows + Microsoft Excel are required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DoubleBookingUiInstallError("pywin32 is required") from exc

    excel = workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        try:
            vbproject = workbook.VBProject
            _ = vbproject.VBComponents.Count
        except Exception as exc:
            raise DoubleBookingUiInstallError(
                "Excel blocked programmatic access to the VBA project. Enable Trust access to the VBA project object model."
            ) from exc

        install_double_booking_review(vbproject)
        install_double_booking_daily_projection(vbproject)
        _replace_form_code(vbproject)
        workbook.Save()
        count = int(vbproject.VBComponents.Count)
        return (
            count,
            _component_present(vbproject, DOUBLE_BOOKING_MODULE_NAME),
            _component_present(vbproject, DOUBLE_BOOKING_FORM_NAME),
            _component_present(vbproject, DOUBLE_BOOKING_PROJECTION_MODULE_NAME),
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Install the double-booking review UI into a fresh full-app preview copy.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.input.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: input workbook not found: {source}")
        return 2
    if source == output:
        print("SAFETY STOP: output must be different from input")
        return 2
    if output.exists() and not args.overwrite:
        print(f"SAFETY STOP: output already exists: {output}")
        return 2

    source_hash = file_sha256(source)
    output_created = False
    try:
        source_contract = assert_application_contract(source)
        output.parent.mkdir(parents=True, exist_ok=True)
        if output.exists():
            output.unlink()
        shutil.copy2(source, output)
        output_created = True
        component_count, module_present, form_present, projection_present = install_ui(output)
        if not module_present or not form_present or not projection_present:
            raise DoubleBookingUiInstallError(
                "Double-booking VBA components were not confirmed after installation"
            )
        output_contract = assert_application_contract(output)
        source_unchanged = file_sha256(source) == source_hash
        if not source_unchanged:
            raise DoubleBookingUiInstallError("Input workbook changed during UI installation")
    except Exception as exc:
        if output_created and output.exists():
            try:
                output.unlink()
            except OSError:
                pass
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DOUBLE BOOKING REVIEW UI PREVIEW OK")
    print(f"Input: {source}")
    print(f"Preview: {output}")
    print("Source unchanged: True")
    print(f"Application contract source: OK ({len(source_contract.component_names)} components)")
    print(f"Application contract output: OK ({len(output_contract.component_names)} components)")
    print(f"Output VBA component count: {component_count}")
    print(f"Installed module: {DOUBLE_BOOKING_MODULE_NAME}")
    print(f"Installed form: {DOUBLE_BOOKING_FORM_NAME}")
    print(f"Installed projection module: {DOUBLE_BOOKING_PROJECTION_MODULE_NAME}")
    print("Macro: ShowDoubleBookingReview")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
