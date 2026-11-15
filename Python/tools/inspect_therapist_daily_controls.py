from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _safe(label: str, getter) -> str:
    try:
        value = getter()
    except Exception as exc:
        return f"<{label} unavailable: {exc}>"
    if value is None:
        return ""
    return str(value).strip()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inspect native Excel controls/shapes on THERAPIST_DAILY without modifying the workbook."
    )
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    path = args.workbook.resolve()
    if not path.exists():
        print(f"SAFETY STOP: workbook not found: {path}")
        return 2
    if sys.platform != "win32":
        print("SAFETY STOP: Windows Excel is required")
        return 2

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 is required")
        return 2

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        ws = workbook.Worksheets("THERAPIST_DAILY")

        print("THERAPIST_DAILY CONTROL INVENTORY")
        print(f"Workbook: {path}")
        print(f"Shapes: {int(ws.Shapes.Count)}")
        for index in range(1, int(ws.Shapes.Count) + 1):
            shape = ws.Shapes.Item(index)
            name = _safe("name", lambda: shape.Name)
            shape_type = _safe("type", lambda: shape.Type)
            on_action = _safe("on_action", lambda: shape.OnAction)
            text2 = _safe("textframe2", lambda: shape.TextFrame2.TextRange.Text)
            text1 = _safe("textframe", lambda: shape.TextFrame.Characters().Text)
            alt = _safe("alternative_text", lambda: shape.AlternativeText)
            print(
                f"SHAPE {index}: name={name!r} type={shape_type!r} "
                f"text2={text2!r} text={text1!r} alt={alt!r} on_action={on_action!r}"
            )

        try:
            ole_count = int(ws.OLEObjects().Count)
        except Exception:
            ole_count = 0
        print(f"OLEObjects: {ole_count}")
        for index in range(1, ole_count + 1):
            obj = ws.OLEObjects().Item(index)
            name = _safe("name", lambda: obj.Name)
            prog_id = _safe("prog_id", lambda: obj.progID)
            caption = _safe("caption", lambda: obj.Object.Caption)
            print(
                f"OLE {index}: name={name!r} prog_id={prog_id!r} caption={caption!r}"
            )

        try:
            buttons = ws.Buttons()
            button_count = int(buttons.Count)
        except Exception:
            buttons = None
            button_count = 0
        print(f"Buttons: {button_count}")
        if buttons is not None:
            for index in range(1, button_count + 1):
                button = buttons.Item(index)
                name = _safe("name", lambda: button.Name)
                caption = _safe("caption", lambda: button.Caption)
                on_action = _safe("on_action", lambda: button.OnAction)
                print(
                    f"BUTTON {index}: name={name!r} caption={caption!r} on_action={on_action!r}"
                )

        return 0
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
