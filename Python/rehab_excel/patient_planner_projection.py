from __future__ import annotations

from pathlib import Path
import sys

from .outpatient_presentation import OUTPATIENT_LIGHT_BLUE_RGB


class PatientPlannerProjectionError(RuntimeError):
    """Raised when PATIENT_PLANNER identity projection cannot be refreshed safely."""


def _excel_rgb(hex_rgb: str) -> int:
    value = hex_rgb.strip().lstrip("#")
    if len(value) != 6:
        raise ValueError("RGB value must be six hexadecimal digits")
    r = int(value[0:2], 16)
    g = int(value[2:4], 16)
    b = int(value[4:6], 16)
    return r + (g << 8) + (b << 16)


def planner_identity_formulas(row: int) -> tuple[str, str, str, str, str]:
    if row < 2:
        raise ValueError("PATIENT_PLANNER data row must be >= 2")
    return (
        f'=IF(PATIENTS!C{row}="","",PATIENTS!A{row})',
        f'=IF(PATIENTS!C{row}="","",PATIENTS!B{row})',
        f'=IF(PATIENTS!C{row}="","",PATIENTS!C{row})',
        f'=IF(PATIENTS!C{row}="","",PATIENTS!D{row})',
        f'=IF(PATIENTS!C{row}="","",PATIENTS!E{row})',
    )


def is_outpatient_type(value: object) -> bool:
    return str(value or "").strip().casefold() in {
        "εξωτερικός".casefold(),
        "outpatient",
    }


def refresh_patient_planner_projection_in_place(workbook_path: str | Path) -> int:
    """Refresh PATIENT_PLANNER identity/status formulas on a working .xlsm copy.

    PATIENTS remains authoritative. Columns A:E of PATIENT_PLANNER mirror the
    registry row-for-row, but blank source cells stay blank instead of showing
    Excel's numeric zero. Outpatients keep an empty infectious value, preserve
    their operational status, and receive the shared outpatient light-blue
    marker on the patient-name cell only.
    """

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise PatientPlannerProjectionError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise PatientPlannerProjectionError("PATIENT_PLANNER refresh requires .xlsm")

    if sys.platform != "win32":
        raise PatientPlannerProjectionError(
            "PATIENT_PLANNER refresh requires Windows with Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise PatientPlannerProjectionError("pywin32 is required") from exc

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)

        try:
            patients = workbook.Worksheets("PATIENTS")
            planner = workbook.Worksheets("PATIENT_PLANNER")
        except Exception as exc:
            raise PatientPlannerProjectionError(
                "Workbook requires PATIENTS and PATIENT_PLANNER sheets"
            ) from exc

        headers: dict[str, int] = {}
        last_header_col = max(7, int(patients.UsedRange.Columns.Count))
        for col in range(1, last_header_col + 1):
            text = str(patients.Cells(1, col).Value2 or "").strip()
            if text:
                headers[text] = col

        type_col = None
        for name in ("PatientType", "ΤύποςΑσθενή", "Τύπος Ασθενή"):
            if name in headers:
                type_col = headers[name]
                break
        if type_col is None:
            raise PatientPlannerProjectionError(
                "PATIENTS has no patient-type column"
            )

        last_patient_row = max(
            int(patients.Cells(patients.Rows.Count, 1).End(-4162).Row),
            int(patients.Cells(patients.Rows.Count, 3).End(-4162).Row),
        )
        last_planner_row = max(
            int(planner.Cells(planner.Rows.Count, 1).End(-4162).Row),
            int(planner.UsedRange.Rows.Count),
        )
        last_row = max(2, last_patient_row, last_planner_row)

        blue = _excel_rgb(OUTPATIENT_LIGHT_BLUE_RGB)
        for row in range(2, last_row + 1):
            formulas = planner_identity_formulas(row)
            for col, formula in enumerate(formulas, start=1):
                planner.Cells(row, col).Formula = formula

            name_cell = planner.Cells(row, 3)
            try:
                if int(name_cell.Interior.Color) == blue:
                    name_cell.Interior.Pattern = -4142
            except Exception:
                pass

            patient_type = patients.Cells(row, type_col).Value2
            if is_outpatient_type(patient_type):
                name_cell.Interior.Color = blue

        # The next registration stages depend on the formulas themselves, not on
        # freshly calculated cached values. A full-workbook recalculation here is
        # therefore redundant and was the dominant hidden delay on large files.
        workbook.Save()
        return max(0, last_patient_row - 1)
    except PatientPlannerProjectionError:
        raise
    except Exception as exc:
        raise PatientPlannerProjectionError(
            f"Excel PATIENT_PLANNER refresh failed: {exc}"
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
