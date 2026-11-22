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
from rehab_excel.vba_form_designer_audit import (
    audit_userform_designers, assert_userform_designers_preserved,
)


class SaveComparisonError(RuntimeError):
    pass


def _save_copy(copy: Path, *, add_journal: bool, excel_factory, designer_readback=False):
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
        baseline_designer = (
            audit_userform_designers(book, expected_form_names=forms)
            if designer_readback else ()
        )
        if add_journal:
            if any(str(book.Worksheets(i).Name) == JOURNAL_SHEET
                   for i in range(1, int(book.Worksheets.Count) + 1)):
                raise SaveComparisonError("Journal already present in source")
            ensure_daily_input_journal_sheet(book)
        else:
            book.Save()
        book.Close(SaveChanges=False)
        book = None
        saved_designer = ()
        if designer_readback:
            # Reopen only the disposable copy. Never trigger macros, events or
            # another save. Compare each workbook against its own baseline.
            book = app.Workbooks.Open(
                str(copy), UpdateLinks=0, ReadOnly=True,
                AddToMru=False, IgnoreReadOnlyRecommended=True,
            )
            if not bool(book.ReadOnly):
                raise SaveComparisonError("Designer verification reopen was not read-only")
            saved_designer = audit_userform_designers(
                book, expected_form_names=forms,
            )
            assert_userform_designers_preserved(baseline_designer, saved_designer)
            book.Close(SaveChanges=False)
            book = None
        return forms, saved_designer
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


def compare_noedit_and_journal_saves(
    source_path: Path, *, excel_factory=None, control_only=False,
    designer_readback=False,
):
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
            control_forms, control_designers = _save_copy(
                noedit, add_journal=False, excel_factory=excel_factory,
                designer_readback=designer_readback,
            )
            stage = "save second control without edits" if control_only else "save journal trial"
            journal_forms, journal_designers = _save_copy(
                journal, add_journal=not control_only, excel_factory=excel_factory,
                designer_readback=designer_readback,
            )
            if sorted(x.casefold() for x in control_forms) != sorted(x.casefold() for x in journal_forms):
                raise SaveComparisonError("UserForm inventory differs between controls")
            if designer_readback:
                stage = "compare cross-save UserForm designer semantics"
                assert_userform_designers_preserved(
                    control_designers, journal_designers,
                )
            stage = "compare embedded OLE streams"
            control_ole = form_storage_evidence(noedit, control_forms)
            journal_ole = form_storage_evidence(journal, control_forms)
            changes = compare_form_storage_evidence(control_ole, journal_ole)
            categories = stream_change_categories(control_ole, journal_ole)
        stage = "verify original unchanged"
        if _file_sha256(source) != before:
            raise SaveComparisonError("Source changed during comparative probe")
        return {"form_count": len(control_forms), "differences": changes, "categories": categories,
                "designer_forms_verified": len(control_designers) if designer_readback else None,
                "comparison": "NO-EDIT vs NO-EDIT" if control_only else "NO-EDIT vs JOURNAL"}
    except Exception as exc:
        detail = " ".join(str(exc).replace(str(source), "<SOURCE>").split())[:300]
        raise SaveComparisonError(f"Stage={stage}; cause={type(exc).__name__}; detail={detail}") from exc


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Diagnostic: Excel no-edit vs journal save, temporary copies only")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--control-only", action="store_true",
                        help="Compare two independent no-edit saves; never install journal")
    parser.add_argument("--designer-readback", action="store_true",
                        help="Verify both disposable UserForm designer trees; diagnostic only")
    args = parser.parse_args(argv)
    try:
        result = compare_noedit_and_journal_saves(
            args.source, control_only=args.control_only,
            designer_readback=args.designer_readback,
        )
    except Exception as exc:
        print("SAFETY STOP: NO-EDIT / JOURNAL COMPARISON FAILED")
        print("DIAGNOSTIC: " + " ".join(str(exc).split())[:500])
        return 2
    print("EXCEL INDEPENDENT SAVE COMPARISON COMPLETED")
    print("Comparison: " + result["comparison"])
    print(f"UserForms compared: {result['form_count']}")
    if result["designer_forms_verified"] is not None:
        print(f"Designer structural readback: {result['designer_forms_verified']} forms matched")
        print("Designer comparisons: baseline-to-reopen each copy + cross-save matched")
    for name, count in result["differences"].items():
        print(f"{name}={count}")
    for name, count in result["categories"].items():
        print(f"{name}={count}")
    print("Source untouched: YES")
    print("Both temporary copies removed; no operational workbook created")
    print("DIAGNOSTIC ONLY: a zero or nonzero comparison is not an integrity approval")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
