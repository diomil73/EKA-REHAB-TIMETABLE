from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]


def _member_info(path: Path, member: str) -> tuple[bool, int | None, str | None]:
    with ZipFile(path) as archive:
        try:
            data = archive.read(member)
        except KeyError:
            return False, None, None
    return True, len(data), sha256(data).hexdigest()


def _excel_info(path: Path) -> tuple[bool, int]:
    import win32com.client  # type: ignore[import-not-found]

    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        return bool(workbook.HasVBProject), int(workbook.FileFormat)
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
    parser = argparse.ArgumentParser(description="Diagnose VBA preservation in an .xlsm preview without changing either workbook.")
    parser.add_argument("--source", type=Path, default=REPO_ROOT / "Rehab_Center_System_v27_1.xlsm")
    parser.add_argument("--preview", type=Path, default=REPO_ROOT / "Excel" / "previews" / "PSYCHOLOGY_SCHEMA_PREVIEW.xlsm")
    args = parser.parse_args()

    source = args.source.resolve()
    preview = args.preview.resolve()

    if sys.platform != "win32":
        print("SAFETY STOP: this diagnostic requires Windows with Microsoft Excel.")
        return 2
    for label, path in (("SOURCE", source), ("PREVIEW", preview)):
        if not path.exists():
            print(f"SAFETY STOP: {label.lower()} workbook not found: {path}")
            return 2

    try:
        import win32com.client  # noqa: F401  # type: ignore[import-not-found]
    except ImportError:
        print("SAFETY STOP: pywin32 is required (pip install pywin32).")
        return 2

    src_present, src_size, src_hash = _member_info(source, "xl/vbaProject.bin")
    out_present, out_size, out_hash = _member_info(preview, "xl/vbaProject.bin")
    src_has_project, src_format = _excel_info(source)
    out_has_project, out_format = _excel_info(preview)

    print("VBA PRESERVATION DIAGNOSTIC")
    print(f"SOURCE zip vbaProject.bin: present={src_present}, size={src_size}")
    print(f"PREVIEW zip vbaProject.bin: present={out_present}, size={out_size}")
    print(f"VBA binary identical: {src_hash is not None and src_hash == out_hash}")
    print(f"SOURCE Excel HasVBProject: {src_has_project}, FileFormat={src_format}")
    print(f"PREVIEW Excel HasVBProject: {out_has_project}, FileFormat={out_format}")

    logically_preserved = src_present and out_present and src_has_project and out_has_project
    print(f"VBA project logically preserved: {logically_preserved}")
    return 0 if logically_preserved else 3


if __name__ == "__main__":
    raise SystemExit(main())
