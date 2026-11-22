"""Compare UserForm OLE streams after no-edit Excel save vs journal-sheet save.

Diagnostic only. Uses two separate fresh disposable XLSM copies, never the
original Excel workbook. Neither result is approved for clinical use.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[2]
for item in (ROOT, ROOT / "Python"):
    if str(item) not in sys.path:
        sys.path.insert(0, str(item))

from Python.tools.prepare_native_probe_copy import create_probe_copy
from Python.tools.provision_daily_input_journal_native import _file_sha256
from rehab_excel.daily_input_journal import JOURNAL_SHEET
from rehab_excel.daily_input_journal_sheet import ensure_daily_input_journal_sheet
from rehab_excel.vba_form_storage import (
    form_storage_evidence, compare_form_storage_evidence, stream_change_categories,
)
from rehab_excel.vba_semantic_audit import audit_vba_components, FORM_TYPE


class SaveComparisonError(RuntimeError):
    pass


def _save_copy(copy: Path, *, add_journal: bool, excel_factory):
    app = None
    book = None
    try:
        app = excel_factory()
        app.AutomationSecurity = 3
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        app.AskToUpdateLinks = False
        book = app.Workbooks.Open(
            str(copy), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(book.ReadOnly):
            raise SaveComparisonError("Disposable copy unexpectedly read-only")
        if bool(book.VBASigned):
            raise SaveComparisonError("Refusing signed VBA project")
        components = audit_vba_components(book)
        forms = tuple(item.name for item in components if item.component_type == FORM_TYPE)
        if not forms:
            raise SaveComparisonError("No UserForms available")
        if add_journal:
            if any(str(book.Worksheets(i).Name) == JOURNAL_SHEET
                   for i in range(1, int(book.Worksheets.Count) + 1)):
                raise SaveComparisonError("Journal already present in source")
            ensure_daily_input_journal_sheet(book)
        else:
            book.Save()
        book.Close(SaveChanges=False)
        book = None
        return forms
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


def compare_noedit_and_journal_saves(source_path: Path, *, excel_factory=None):
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise SaveComparisonError("Existing demo XLSM source required")
    if excel_factory is None:
        if sys.platform != "win32":
            raise SaveComparisonError("Windows native Excel required")
        try:
            import win32com.client
        except ImportError as exc:
            raise SaveComparisonError("pywin32 is required") from exc
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")
    before = _file_sha256(source)
    stage = "create independent disposable copies"
    try:
        with TemporaryDirectory(prefix="diomil_ab_ole_") as folder:
            noedit = Path(folder) / "noedit.xlsm"
            journal = Path(folder) / "journal.xlsm"
            create_probe_copy(source, noedit)
            create_probe_copy(source, journal)
            stage = "save no-edit control"
            control_forms = _save_copy(noedit, add_journal=False, excel_factory=excel_factory)
            stage = "save journal trial"
            journal_forms = _save_copy(journal, add_journal=True, excel_factory=excel_factory)
            if sorted(x.casefold() for x in control_forms) != sorted(x.casefold() for x in journal_forms):
                raise SaveComparisonError("UserForm inventory differs between controls")
            stage = "compare embedded OLE streams"
            control_ole = form_storage_evidence(noedit, control_forms)
            journal_ole = form_storage_evidence(journal, control_forms)
            changes = compare_form_storage_evidence(control_ole, journal_ole)
            categories = stream_change_categories(control_ole, journal_ole)
        stage = "verify original unchanged"
        if _file_sha256(source) != before:
            raise SaveComparisonError("Source changed during comparative probe")
        return {"form_count": len(control_forms), "differences": changes, "categories": categories}
    except Exception as exc:
        detail = " ".join(str(exc).replace(str(source), "<SOURCE>").split())[:300]
        raise SaveComparisonError(f"Stage={stage}; cause={type(exc).__name__}; detail={detail}") from exc


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Diagnostic: Excel no-edit vs journal save, temporary copies only")
    parser.add_argument("--source", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = compare_noedit_and_journal_saves(args.source)
    except Exception as exc:
        print("SAFETY STOP: NO-EDIT / JOURNAL COMPARISON FAILED")
        print("DIAGNOSTIC: " + " ".join(str(exc).split())[:500])
        return 2
    print("EXCEL NO-EDIT VS JOURNAL SAVE COMPARISON COMPLETED")
    print(f"UserForms compared: {result['form_count']}")
    for name, count in result["differences"].items():
        print(f"{name}={count}")
    for name, count in result["categories"].items():
        print(f"{name}={count}")
    print("Source untouched: YES")
    print("Both temporary copies removed; no operational workbook created")
    print("DIAGNOSTIC ONLY: equality does not prove semantic integrity")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
