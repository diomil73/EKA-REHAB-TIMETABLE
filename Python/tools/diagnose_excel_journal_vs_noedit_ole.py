"""Compare UserForm OLE streams after no-edit Excel save vs journal-sheet save.

Diagnostic only. Uses two separate fresh disposable XLSM copies, never the
original Excel workbook. Neither result is approved for clinical use.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[2]
for item in (ROOT, ROOT / "Python"):
    if str(item) not in sys.path:
        sys.path.insert(0, str(item))

from Python.tools.prepare_native_probe_copy import create_probe_copy
from Python.tools.provision_daily_input_journal_native import _file_sha256
from rehab_excel.daily_input_journal import JOURNAL_SHEET
from rehab_excel.daily_input_journal_sheet import (
    ensure_daily_input_journal_sheet, append_native_operation_revision,
    read_native_journal_rows,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation, DailyInputOperationPhase, advance_daily_input_operation,
)
from tools.prepare_daily_input_journal_prepared import (
    prepare_prepared_handoff, _source_date,
)
from tools.verify_daily_input_journal_prepared import verify_prepared_reopen
from rehab_excel.vba_form_storage import (
    form_storage_evidence, compare_form_storage_evidence, stream_change_categories,
    assert_matching_source_relative_change_footprints,
)
from rehab_excel.vba_semantic_audit import audit_vba_components, FORM_TYPE
from rehab_excel.vba_form_designer_audit import (
    audit_userform_designers, assert_userform_designers_preserved,
)


class SaveComparisonError(RuntimeError):
    pass


def _create_fictional_prepared_handoff(copy: Path) -> dict[str, object]:
    """Bind a wholly fictional payload receipt to a disposable saved workbook."""
    from openpyxl import load_workbook

    payload_path = copy.with_name(copy.stem + ".fictional-payload.tsv")
    response_path = copy.with_name(copy.stem + ".fictional-response.json")
    handoff_path = copy.with_name(copy.stem + ".fictional-handoff.json")
    if any(path.exists() for path in (payload_path, response_path, handoff_path)):
        raise SaveComparisonError("Fictional handoff evidence path already exists")

    workbook = None
    try:
        workbook = load_workbook(
            copy, read_only=True, data_only=False, keep_vba=True,
        )
        if "DAILY_INPUT" not in workbook.sheetnames:
            raise SaveComparisonError("Fictional prepared trial requires DAILY_INPUT")
        service_date = _source_date(workbook["DAILY_INPUT"]["B2"].value)
    finally:
        if workbook is not None:
            workbook.close()

    payload_path.write_text("PATCH\tFICTIONAL_ONLY\n", encoding="utf-8")
    receipt = {
        "ok": True,
        "source_unchanged": True,
        "date": service_date,
        "source_sha256_before": _file_sha256(copy),
        "payload_path": str(payload_path.resolve()),
    }
    response_path.write_text(
        json.dumps(receipt, ensure_ascii=True, sort_keys=True),
        encoding="utf-8",
    )
    handoff = prepare_prepared_handoff(
        copy, response_path, payload_path,
        operation_id="FICTIONAL-PREPARED-" + uuid4().hex,
    )
    handoff_path.write_text(
        json.dumps(handoff, ensure_ascii=True, sort_keys=True),
        encoding="utf-8",
    )
    return {
        "handoff": handoff,
        "payload_path": payload_path,
        "handoff_path": handoff_path,
    }


def _fictional_journal_states(*, sequential=False):
    """Fixture-only lifecycle. No real patient identifiers or clinical facts."""
    first = DailyInputOperation(
        operation_id="FICTIONAL-NATIVE-JOURNAL-TRIAL",
        source_sha256="0" * 64,
        service_date="2000-01-03",
    )
    failure = advance_daily_input_operation(
        first, DailyInputOperationPhase.FAILED, failure_code="FICTIONAL_PROBE",
    )
    restored = advance_daily_input_operation(
        failure, DailyInputOperationPhase.ROLLED_BACK, rollback_verified=True,
    )
    if not sequential:
        return first, failure, restored
    # Same journal worksheet: a new operation after verified terminal state.
    second = DailyInputOperation(
        operation_id="FICTIONAL-NATIVE-JOURNAL-TRIAL-02",
        source_sha256="1" * 64,
        service_date="2000-01-04",
    )
    second_failed = advance_daily_input_operation(
        second, DailyInputOperationPhase.FAILED, failure_code="FICTIONAL_PROBE",
    )
    second_restored = advance_daily_input_operation(
        second_failed, DailyInputOperationPhase.ROLLED_BACK, rollback_verified=True,
    )
    return first, failure, restored, second, second_failed, second_restored


def _save_copy(
    copy: Path, *, add_journal: bool, excel_factory,
    designer_readback=False, trial_fictional_revisions=False,
    trial_sequential_operations=False, trial_prepared_handoff=False,
):
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
        staged_rows = ()
        fictional_revision_count = 0
        prepared_handoff = None
        prepared_handoff_verified = False
        if add_journal:
            if any(str(book.Worksheets(i).Name) == JOURNAL_SHEET
                   for i in range(1, int(book.Worksheets.Count) + 1)):
                raise SaveComparisonError("Journal already present in source")
            journal_sheet = ensure_daily_input_journal_sheet(book)
            if trial_fictional_revisions:
                for operation in _fictional_journal_states(
                    sequential=trial_sequential_operations,
                ):
                    if not append_native_operation_revision(book, operation):
                        raise SaveComparisonError("Fictional journal revision was not appended")
                staged_rows = read_native_journal_rows(journal_sheet)
                expected_count = 6 if trial_sequential_operations else 3
                if len(staged_rows) != expected_count:
                    raise SaveComparisonError("Fictional journal staging incomplete")
                fictional_revision_count = len(staged_rows)
                book.Save()

            if trial_prepared_handoff:
                if not trial_sequential_operations or fictional_revision_count != 6:
                    raise SaveComparisonError(
                        "Prepared handoff trial requires two terminal fictional operations"
                    )
                prepared_handoff = _create_fictional_prepared_handoff(copy)
                metadata = prepared_handoff["handoff"]
                prepared_operation = DailyInputOperation(
                    operation_id=metadata["operation_id"],
                    source_sha256=metadata["source_sha256"],
                    service_date=metadata["service_date"],
                )
                if not append_native_operation_revision(book, prepared_operation):
                    raise SaveComparisonError("Prepared journal revision was not appended")
                staged_rows = read_native_journal_rows(journal_sheet)
                if len(staged_rows) != fictional_revision_count + 1:
                    raise SaveComparisonError("Prepared journal row count mismatch")
                book.Save()
        else:
            book.Save()
        book.Close(SaveChanges=False)
        book = None
        if trial_prepared_handoff:
            if prepared_handoff is None:
                raise SaveComparisonError("Prepared handoff evidence was not created")
            verify_prepared_reopen(
                copy, prepared_handoff["handoff"], prepared_handoff["payload_path"],
            )
            prepared_handoff_verified = True
        saved_designer = ()
        if designer_readback or trial_fictional_revisions:
            # Reopen only the disposable copy. Never trigger macros, events or
            # another save. Compare each workbook against its own baseline.
            book = app.Workbooks.Open(
                str(copy), UpdateLinks=0, ReadOnly=True,
                AddToMru=False, IgnoreReadOnlyRecommended=True,
            )
            if not bool(book.ReadOnly):
                raise SaveComparisonError("Designer verification reopen was not read-only")
            if designer_readback:
                saved_designer = audit_userform_designers(
                    book, expected_form_names=forms,
                )
                assert_userform_designers_preserved(baseline_designer, saved_designer)
            if trial_fictional_revisions:
                saved_journal = book.Worksheets(JOURNAL_SHEET)
                if int(saved_journal.Visible) != 2 or not bool(saved_journal.ProtectContents):
                    raise SaveComparisonError("Reopened fictional journal not protected/veryHidden")
                if read_native_journal_rows(saved_journal) != staged_rows:
                    raise SaveComparisonError("Fictional journal changed on native reopen")
            book.Close(SaveChanges=False)
            book = None
        return (
            forms, saved_designer, fictional_revision_count,
            prepared_handoff_verified, len(staged_rows),
        )
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
    designer_readback=False, trial_fictional_revisions=False,
    trial_sequential_operations=False, trial_prepared_handoff=False,
):
    if trial_prepared_handoff and (
        control_only or not designer_readback
        or not trial_fictional_revisions or not trial_sequential_operations
    ):
        raise SaveComparisonError(
            "Prepared handoff trial requires sequential fictional revisions and designer readback"
        )
    if trial_sequential_operations and not trial_fictional_revisions:
        raise SaveComparisonError("Sequential trial requires fictional revisions")
    if trial_fictional_revisions and (control_only or not designer_readback):
        raise SaveComparisonError("Fictional journal trial requires journal and designer readback")
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
            control_forms, control_designers, _, _, _ = _save_copy(
                noedit, add_journal=False, excel_factory=excel_factory,
                designer_readback=designer_readback,
            )
            stage = "save second control without edits" if control_only else "save journal trial"
            (
                journal_forms, journal_designers, verified_trial_rows,
                prepared_handoff_verified, prepared_journal_rows,
            ) = _save_copy(
                journal, add_journal=not control_only, excel_factory=excel_factory,
                designer_readback=designer_readback,
                trial_fictional_revisions=trial_fictional_revisions,
                trial_sequential_operations=trial_sequential_operations,
                trial_prepared_handoff=trial_prepared_handoff,
            )
            if sorted(x.casefold() for x in control_forms) != sorted(x.casefold() for x in journal_forms):
                raise SaveComparisonError("UserForm inventory differs between controls")
            if designer_readback:
                stage = "compare cross-save UserForm designer semantics"
                assert_userform_designers_preserved(
                    control_designers, journal_designers,
                )
            stage = "compare embedded OLE streams"
            source_ole = form_storage_evidence(source, control_forms)
            control_ole = form_storage_evidence(noedit, control_forms)
            journal_ole = form_storage_evidence(journal, control_forms)
            changes = compare_form_storage_evidence(control_ole, journal_ole)
            categories = stream_change_categories(control_ole, journal_ole)
            stage = "verify source-relative OLE change footprints"
            matching_changed_streams = assert_matching_source_relative_change_footprints(
                source_ole, control_ole, journal_ole,
            )
        stage = "verify original unchanged"
        if _file_sha256(source) != before:
            raise SaveComparisonError("Source changed during comparative probe")
        return {"form_count": len(control_forms), "differences": changes, "categories": categories,
                "designer_forms_verified": len(control_designers) if designer_readback else None,
                "comparison": "NO-EDIT vs NO-EDIT" if control_only else "NO-EDIT vs JOURNAL",
                "source_relative_footprint_streams": matching_changed_streams,
                "fictional_revisions_verified": verified_trial_rows,
                "prepared_handoff_verified": prepared_handoff_verified,
                "prepared_journal_rows": prepared_journal_rows}
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
    parser.add_argument("--trial-fictional-revisions", action="store_true",
                        help="On journal disposable copy only, stage/reopen 3 synthetic rollback revisions")
    parser.add_argument("--trial-sequential-operations", action="store_true",
                        help="With fictional revisions, additionally stage a second unique terminal operation")
    parser.add_argument("--trial-prepared-handoff", action="store_true",
                        help="After two fictional terminal operations, append one source-bound PREPARED row and independently verify after native reopen")
    args = parser.parse_args(argv)
    try:
        result = compare_noedit_and_journal_saves(
            args.source, control_only=args.control_only,
            designer_readback=args.designer_readback,
            trial_fictional_revisions=args.trial_fictional_revisions,
            trial_sequential_operations=args.trial_sequential_operations,
            trial_prepared_handoff=args.trial_prepared_handoff,
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
    print("Source-relative OLE change footprint matched: YES")
    print(f"Source-relative changed stream identities matched: {result['source_relative_footprint_streams']}")
    print("Not binary/resource preservation approval")
    if result["fictional_revisions_verified"]:
        print(f"Fictional journal revisions saved/reopened: {result['fictional_revisions_verified']}")
    if result["prepared_handoff_verified"]:
        print("Fictional PREPARED handoff saved/reopened: 1")
        print(f"Total journal revisions after readback: {result['prepared_journal_rows']}")
    if result["fictional_revisions_verified"] or result["prepared_handoff_verified"]:
        print("Synthetic journal only: no operational Apply/Restore executed")
    print("Source untouched: YES")
    print("Both temporary copies removed; no operational workbook created")
    print("DIAGNOSTIC ONLY: a zero or nonzero comparison is not an integrity approval")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
