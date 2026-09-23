from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SOURCE = REPO_ROOT / "Rehab_Center_System_v27_1.xlsm"
PREVIEW = REPO_ROOT / "Excel" / "previews" / "PSYCHOLOGY_SCHEMA_PREVIEW.xlsm"


def _inspect(path: Path) -> tuple[bool, list[tuple[str, int, int]] | None, str | None]:
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

        has_vba = bool(workbook.HasVBProject)
        if not has_vba:
            return False, [], None

        try:
            vbproject = workbook.VBProject
            components = []
            for i in range(1, vbproject.VBComponents.Count + 1):
                component = vbproject.VBComponents.Item(i)
                try:
                    lines = int(component.CodeModule.CountOfLines)
                except Exception:
                    lines = -1
                components.append((str(component.Name), int(component.Type), lines))
            components.sort(key=lambda item: (item[1], item[0].casefold()))
            return True, components, None
        except Exception as exc:
            return True, None, str(exc)
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
    print("VBA COMPONENT DIAGNOSTIC")
    if not SOURCE.exists():
        print(f"SOURCE missing: {SOURCE}")
        return 2
    if not PREVIEW.exists():
        print(f"PREVIEW missing: {PREVIEW}")
        return 2

    source_has, source_components, source_error = _inspect(SOURCE)
    preview_has, preview_components, preview_error = _inspect(PREVIEW)

    print(f"SOURCE HasVBProject: {source_has}")
    print(f"PREVIEW HasVBProject: {preview_has}")

    if source_components is None or preview_components is None:
        print("VBProject component inspection unavailable.")
        if source_error:
            print(f"SOURCE VBProject access error: {source_error}")
        if preview_error:
            print(f"PREVIEW VBProject access error: {preview_error}")
        print(
            "NOTE: Excel may have 'Trust access to the VBA project object model' disabled."
        )
        return 3

    print(f"SOURCE component count: {len(source_components)}")
    for item in source_components:
        print(f"  SOURCE {item[0]} | type={item[1]} | lines={item[2]}")

    print(f"PREVIEW component count: {len(preview_components)}")
    for item in preview_components:
        print(f"  PREVIEW {item[0]} | type={item[1]} | lines={item[2]}")

    same = source_components == preview_components
    print(f"VBA component structure identical: {same}")
    return 0 if same else 4


if __name__ == "__main__":
    raise SystemExit(main())
