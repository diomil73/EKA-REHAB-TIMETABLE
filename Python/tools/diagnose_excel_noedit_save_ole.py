"""Isolated control experiment: save a disposable demo XLSM in Excel without edits.

Does not create a journal, modify VBA, or touch source workbook in Excel.
Only reports aggregate UserForm OLE stream changes; never prints names, data,
paths or content hashes. Always removes its temporary copy.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[2]
for directory in (ROOT, ROOT / "Python"):
    if str(directory) not in sys.path:
        sys.path.insert(0, str(directory))

from Python.tools.prepare_native_probe_copy import create_probe_copy
from Python.tools.provision_daily_input_journal_native import _file_sha256
from rehab_excel.vba_semantic_audit import audit_vba_components, FORM_TYPE
from rehab_excel.vba_form_storage import (form_storage_evidence, compare_form_storage_evidence, stream_change_categories)


class NativeNoEditSaveError(RuntimeError):
    pass


def probe_excel_save_without_edits(source_path: Path, *, excel_factory=None) -> dict:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.lower() != ".xlsm":
        raise NativeNoEditSaveError("Existing demo XLSM source required")
    if excel_factory is None:
        if sys.platform != "win32":
            raise NativeNoEditSaveError("Windows native Excel required")
        import win32com.client
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    before = _file_sha256(source)
    stage = "prepare isolated disposable copy"
    app = None
    book = None
    try:
        with TemporaryDirectory(prefix="diomil_noedit_save_") as directory:
            copy = Path(directory) / "noedit_save_probe.xlsm"
            create_probe_copy(source, copy)
            stage = "start isolated Excel"
            app = excel_factory()
            app.AutomationSecurity = 3
            app.Visible = False
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.AskToUpdateLinks = False

            stage = "open disposable workbook"
            book = app.Workbooks.Open(
                str(copy), UpdateLinks=0, ReadOnly=False,
                AddToMru=False, IgnoreReadOnlyRecommended=True,
            )
            if bool(book.ReadOnly):
                raise NativeNoEditSaveError("Disposable copy opened read-only")
            stage = "inventory existing UserForm components"
            components = audit_vba_components(book)
            form_names = tuple(item.name for item in components
                               if item.component_type == FORM_TYPE)
            if not form_names:
                raise NativeNoEditSaveError("No UserForms to compare")
            stage = "snapshot original embedded UserForm streams"
            original_streams = form_storage_evidence(source, form_names)
            stage = "save disposable workbook without edits"
            book.Save()
            book.Close(SaveChanges=False)
            book = None
            stage = "compare saved embedded UserForm streams"
            saved_streams = form_storage_evidence(copy, form_names)
            differences = compare_form_storage_evidence(original_streams, saved_streams)
            categories = stream_change_categories(original_streams, saved_streams)
            stage = "verify no source change"
            if _file_sha256(source) != before:
                raise NativeNoEditSaveError("Source changed unexpectedly")
            return {"forms": len(form_names), "differences": differences, "categories": categories}
    except Exception as exc:
        detail = " ".join(str(exc).replace(str(source), "<SOURCE>").split())[:350]
        raise NativeNoEditSaveError(
            f"Stage={stage}; cause={type(exc).__name__}; detail={detail}"
        ) from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="No-edit native Excel save control, disposable copy only")
    parser.add_argument("--source", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = probe_excel_save_without_edits(args.source)
    except Exception as exc:
        print("SAFETY STOP: NO-EDIT NATIVE CONTROL FAILED")
        print("DIAGNOSTIC: " + " ".join(str(exc).split())[:500])
        return 2
    print("NO-EDIT NATIVE EXCEL SAVE CONTROL COMPLETED")
    print("UserForms checked: " + str(result["forms"]))
    print("Embedded stream aggregate differences:")
    for category, count in result["differences"].items():
        print(f"{category}={count}")
    print("Known stream categories:")
    for category, count in result["categories"].items():
        print(f"{category}={count}")
    print("Source unchanged: YES")
    print("No journal or VBA modifications requested; disposable copy deleted")
    print("DIAGNOSTIC ONLY: This is not an integrity approval")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
