from __future__ import annotations

import argparse
from pathlib import Path
import sys


def _safe_text(shape) -> str:
    try:
        return str(shape.TextFrame2.TextRange.Text).replace("\r", " ").replace("\n", " ").strip()
    except Exception:
        try:
            return str(shape.TextFrame.Characters().Text).replace("\r", " ").replace("\n", " ").strip()
        except Exception:
            return ""


def audit(path: Path) -> int:
    """Enumerate saved controls and their assignments; never click, save or edit."""
    if sys.platform != "win32" or not path.is_file():
        print("SAFETY STOP: Windows Excel and existing workbook required")
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
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        print("H7 THERAPIST_DAILY SHAPE READ-ONLY AUDIT")
        print("Workbook:", workbook.FullName)
        ws = workbook.Worksheets("THERAPIST_DAILY")
        for index in range(1, int(ws.Shapes.Count) + 1):
            shape = ws.Shapes(index)
            label = _safe_text(shape)
            try:
                macro = str(shape.OnAction)
            except Exception:
                macro = ""
            try:
                color = int(shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB)
            except Exception:
                color = None
            print(
                f"SHAPE {index}: name={shape.Name!s} caption={label!r} "
                f"macro={macro!r} left={float(shape.Left):.1f} "
                f"top={float(shape.Top):.1f} width={float(shape.Width):.1f} "
                f"height={float(shape.Height):.1f} visible={bool(shape.Visible)} "
                f"font_rgb={color}"
            )
        print("H7 SHAPE AUDIT COMPLETE; workbook unchanged")
        return 0
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        return 2
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
    parser = argparse.ArgumentParser(description="Read-only audit of THERAPIST_DAILY shape geometry and macro binding")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()
    return audit(args.workbook)


if __name__ == "__main__":
    raise SystemExit(main())
