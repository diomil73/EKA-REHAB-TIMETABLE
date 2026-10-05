from __future__ import annotations

from pathlib import Path
import sys

from .native_excel import excel_rgb


class MasterVisualStyleError(RuntimeError):
    """Raised when the MASTER visual styling pass cannot be applied safely."""


# Accepted visual baseline: stronger header shade with a lighter body shade.
HEADER_COLORS = {
    "identity": excel_rgb(199, 212, 228),
    "fth": excel_rgb(111, 157, 198),
    "robotic": excel_rgb(138, 149, 193),
    "pool": excel_rgb(94, 175, 198),
    "reclined": excel_rgb(157, 180, 204),
    "ergo": excel_rgb(143, 184, 117),
    "logo": excel_rgb(214, 180, 91),
    "efa": excel_rgb(165, 192, 122),
    "psych": excel_rgb(215, 154, 124),
    "afternoon": excel_rgb(181, 162, 203),
    "status": excel_rgb(199, 212, 228),
}

BODY_COLORS = {
    "identity": excel_rgb(247, 247, 247),
    "fth": excel_rgb(217, 234, 247),
    "robotic": excel_rgb(230, 228, 243),
    "pool": excel_rgb(216, 240, 244),
    "reclined": excel_rgb(232, 239, 246),
    "ergo": excel_rgb(228, 240, 220),
    "logo": excel_rgb(252, 240, 207),
    "efa": excel_rgb(237, 244, 227),
    "psych": excel_rgb(250, 229, 219),
    "afternoon": excel_rgb(240, 234, 247),
    "status": excel_rgb(247, 247, 247),
}

INFECTIOUS_YELLOW = excel_rgb(255, 216, 77)
ROW_SEPARATOR_GRAY = excel_rgb(166, 166, 166)
SOFT_VERTICAL_GRAY = excel_rgb(230, 230, 230)
HEADER_SEPARATOR_GRAY = excel_rgb(140, 140, 140)
CLINIC_BANNER_FILL = excel_rgb(242, 242, 242)
BLACK = excel_rgb(0, 0, 0)
SECOND_CLINIC_LABEL = "Β' ΚΛΙΝΙΚΗ"


def _header_row(ws) -> int:
    for row in range(1, 11):
        room = str(ws.Cells(row, 2).Value or "").strip().casefold()
        patient = str(ws.Cells(row, 3).Value or "").strip().casefold()
        if room == "θάλαμος".casefold() and patient == "ασθενής".casefold():
            return row
    raise MasterVisualStyleError("MASTER_SCHEDULE header row was not found")


def _last_patient_row(ws, header_row: int) -> int:
    try:
        room_row = int(ws.Cells(ws.Rows.Count, 2).End(-4162).Row)  # xlUp
        patient_row = int(ws.Cells(ws.Rows.Count, 3).End(-4162).Row)
        return max(header_row, room_row, patient_row)
    except Exception:
        return header_row


def _set_fill(ws, address: str, color: int) -> None:
    ws.Range(address).Interior.Color = color


def _is_infectious(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    text = str(value or "").strip().casefold()
    return text in {
        "ν".casefold(),
        "ναι".casefold(),
        "yes",
        "true",
        "1",
        "1.0",
        "-1",
    }


def _clinic_prefix(room: object) -> str | None:
    text = str(room or "").strip().casefold()
    if not text:
        return None
    first = text[0]
    if first in {"a", "α"}:
        return "A"
    if first in {"b", "β"}:
        return "B"
    return None


def _row_is_blank(ws, row: int) -> bool:
    for col in range(2, 14):
        value = ws.Cells(row, col).Value
        if str(value or "").strip():
            return False
    return True


def _remove_internal_blank_rows(ws, header_row: int, last_row: int) -> int:
    """Remove accidental fully-empty rows from inside the patient table."""

    for row in range(last_row, header_row, -1):
        if _row_is_blank(ws, row):
            has_patient_below = any(
                str(ws.Cells(candidate, 3).Value or "").strip()
                for candidate in range(row + 1, last_row + 1)
            )
            has_patient_above = any(
                str(ws.Cells(candidate, 3).Value or "").strip()
                for candidate in range(header_row + 1, row)
            )
            if has_patient_above and has_patient_below:
                ws.Rows(row).Delete()
                last_row -= 1
    return last_row


def _ensure_second_clinic_banner(ws, header_row: int, last_row: int) -> int:
    """Insert one B-clinic banner between the A15 and B01 patient groups."""

    seen_a = False
    first_b_row: int | None = None
    for row in range(header_row + 1, last_row + 1):
        patient = str(ws.Cells(row, 3).Value or "").strip()
        if not patient:
            continue
        clinic = _clinic_prefix(ws.Cells(row, 2).Value)
        if clinic == "A":
            seen_a = True
        elif clinic == "B" and seen_a:
            first_b_row = row
            break

    if first_b_row is None:
        return last_row

    previous_label = str(ws.Cells(first_b_row - 1, 2).Value or "").strip()
    if previous_label.casefold() == SECOND_CLINIC_LABEL.casefold():
        return last_row

    ws.Rows(first_b_row).Insert()
    banner = ws.Range(f"B{first_b_row}:M{first_b_row}")
    banner.ClearContents()
    banner.Interior.Color = CLINIC_BANNER_FILL
    banner.Font.Color = BLACK
    banner.Font.Bold = True
    banner.HorizontalAlignment = 7  # xlCenterAcrossSelection
    banner.VerticalAlignment = -4108  # xlCenter
    ws.Cells(first_b_row, 2).Value = SECOND_CLINIC_LABEL
    ws.Rows(first_b_row).RowHeight = 24

    top = banner.Borders(8)  # xlEdgeTop
    top.LineStyle = 1
    top.Weight = 2
    top.Color = ROW_SEPARATOR_GRAY
    bottom = banner.Borders(9)  # xlEdgeBottom
    bottom.LineStyle = 1
    bottom.Weight = 2
    bottom.Color = ROW_SEPARATOR_GRAY
    return last_row + 1


def _clear_master_freeze(workbook, ws) -> None:
    """The current MASTER design no longer keeps identity columns frozen."""

    try:
        ws.Activate()
        window = workbook.Application.ActiveWindow
        window.FreezePanes = False
        window.SplitRow = 0
        window.SplitColumn = 0
    except Exception:
        pass


def _make_app_buttons_readable(workbook) -> None:
    """Force black text on every EKA toolbar/footer button."""

    try:
        sheet_count = int(workbook.Worksheets.Count)
    except Exception:
        return

    for sheet_index in range(1, sheet_count + 1):
        ws = workbook.Worksheets(sheet_index)
        try:
            shape_count = int(ws.Shapes.Count)
        except Exception:
            continue
        for shape_index in range(1, shape_count + 1):
            try:
                shape = ws.Shapes(shape_index)
                if not str(shape.Name or "").startswith("ekaToolbar_"):
                    continue
                shape.TextFrame2.TextRange.Font.Fill.Visible = -1
                shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BLACK
                try:
                    shape.TextFrame.Characters().Font.Color = BLACK
                except Exception:
                    pass
            except Exception:
                continue


def _apply_master_visual_style(ws) -> None:
    header_row = _header_row(ws)
    last_row = _last_patient_row(ws, header_row)
    last_row = _remove_internal_blank_rows(ws, header_row, last_row)
    last_row = _ensure_second_clinic_banner(ws, header_row, last_row)

    expected = {
        10: "ΕΦΑ",
        11: "ΨΥΧΟΛΟΓΟΙ",
        12: "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ",
        13: "Κατάσταση",
    }
    for col, label in expected.items():
        actual = str(ws.Cells(header_row, col).Value or "").strip()
        if actual.casefold() != label.casefold():
            raise MasterVisualStyleError(
                f"Unexpected MASTER_SCHEDULE schema at column {col}: {actual!r}"
            )

    # Keep the infectious source formula/data in column A, but remove it from the UI.
    ws.Columns("A").Hidden = True

    widths = {
        "B": 9,
        "C": 24,
        "D": 13,
        "E": 12,
        "F": 12,
        "G": 12,
        "H": 14,
        "I": 14,
        "J": 12,
        "K": 14,
        "L": 17,
        "M": 12,
    }
    for column, width in widths.items():
        ws.Columns(column).ColumnWidth = width

    header_ranges = {
        "B:C": HEADER_COLORS["identity"],
        "D:D": HEADER_COLORS["fth"],
        "E:E": HEADER_COLORS["robotic"],
        "F:F": HEADER_COLORS["pool"],
        "G:G": HEADER_COLORS["reclined"],
        "H:H": HEADER_COLORS["ergo"],
        "I:I": HEADER_COLORS["logo"],
        "J:J": HEADER_COLORS["efa"],
        "K:K": HEADER_COLORS["psych"],
        "L:L": HEADER_COLORS["afternoon"],
        "M:M": HEADER_COLORS["status"],
    }
    for columns, color in header_ranges.items():
        start_col, end_col = columns.split(":")
        _set_fill(ws, f"{start_col}{header_row}:{end_col}{header_row}", color)

    if last_row > header_row:
        body_ranges = {
            "B:C": BODY_COLORS["identity"],
            "D:D": BODY_COLORS["fth"],
            "E:E": BODY_COLORS["robotic"],
            "F:F": BODY_COLORS["pool"],
            "G:G": BODY_COLORS["reclined"],
            "H:H": BODY_COLORS["ergo"],
            "I:I": BODY_COLORS["logo"],
            "J:J": BODY_COLORS["efa"],
            "K:K": BODY_COLORS["psych"],
            "L:L": BODY_COLORS["afternoon"],
            "M:M": BODY_COLORS["status"],
        }
        for columns, color in body_ranges.items():
            start_col, end_col = columns.split(":")
            _set_fill(
                ws,
                f"{start_col}{header_row + 1}:{end_col}{last_row}",
                color,
            )

    header = ws.Range(f"B{header_row}:M{header_row}")
    header.Font.Bold = True
    header.HorizontalAlignment = -4108  # xlCenter
    header.VerticalAlignment = -4108
    header.WrapText = True
    ws.Rows(header_row).RowHeight = 34

    if last_row > header_row:
        body = ws.Range(f"B{header_row + 1}:M{last_row}")
        body.HorizontalAlignment = -4108
        body.VerticalAlignment = -4108
        body.WrapText = True
        body.Font.Size = 9
        ws.Range(f"C{header_row + 1}:C{last_row}").HorizontalAlignment = -4131  # xlLeft

    for row in range(header_row + 1, last_row + 1):
        patient = str(ws.Cells(row, 3).Value or "").strip()
        room = str(ws.Cells(row, 2).Value or "").strip()
        if patient:
            ws.Rows(row).RowHeight = 64

            # Strong visual alert with no visible infectious text column.
            if _is_infectious(ws.Cells(row, 1).Value):
                _set_fill(ws, f"B{row}:C{row}", INFECTIOUS_YELLOW)

            row_range = ws.Range(f"B{row}:M{row}")
            bottom = row_range.Borders(9)  # xlEdgeBottom
            bottom.LineStyle = 1
            bottom.Weight = 2
            bottom.Color = ROW_SEPARATOR_GRAY

            for edge in (7, 10):  # xlEdgeLeft / xlEdgeRight
                border = row_range.Borders(edge)
                border.LineStyle = 1
                border.Weight = 1
                border.Color = SOFT_VERTICAL_GRAY
        elif room.casefold() == SECOND_CLINIC_LABEL.casefold():
            banner = ws.Range(f"B{row}:M{row}")
            banner.Interior.Color = CLINIC_BANNER_FILL
            banner.Font.Color = BLACK
            banner.Font.Bold = True
            banner.HorizontalAlignment = 7
            banner.VerticalAlignment = -4108
            ws.Rows(row).RowHeight = 24
        elif room:
            # Other room/separator rows stay compact and neutral.
            ws.Rows(row).RowHeight = 18

    header_bottom = header.Borders(9)
    header_bottom.LineStyle = 1
    header_bottom.Weight = 2
    header_bottom.Color = HEADER_SEPARATOR_GRAY

    # Keep the existing app-shell bounds, but start the visible UI at column B.
    try:
        ws.ScrollArea = f"B1:M{last_row + 1}"
    except Exception:
        pass


def apply_master_visual_style_in_place(workbook_path: str | Path) -> None:
    """Apply the accepted MASTER visual baseline to a working/preview XLSM."""

    path = Path(workbook_path).resolve()
    if not path.exists():
        raise MasterVisualStyleError(f"Workbook not found: {path}")
    if path.suffix.casefold() != ".xlsm":
        raise MasterVisualStyleError("MASTER visual styling requires .xlsm")
    if sys.platform != "win32":
        raise MasterVisualStyleError(
            "MASTER visual styling requires Windows with Microsoft Excel"
        )

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise MasterVisualStyleError("pywin32 is required") from exc

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
            ws = workbook.Worksheets("MASTER_SCHEDULE")
        except Exception as exc:
            raise MasterVisualStyleError("Workbook has no MASTER_SCHEDULE sheet") from exc

        excel.CalculateFull()
        _apply_master_visual_style(ws)
        _clear_master_freeze(workbook, ws)
        _make_app_buttons_readable(workbook)
        workbook.Save()
    except MasterVisualStyleError:
        raise
    except Exception as exc:
        raise MasterVisualStyleError(f"MASTER visual styling failed: {exc}") from exc
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
