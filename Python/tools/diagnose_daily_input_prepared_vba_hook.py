"""Run the VBA-owned PREPARED append only on two disposable XLSM copies.

The control and trial copies receive the same journal sheet, six fictional
terminal history rows, and VBA module. Only the trial invokes the VBA hook.
No operational Apply/Restore or patient payload is executed.
"""
from __future__ import annotations

import argparse
from datetime import date, datetime
from hashlib import sha256
import json
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from openpyxl import load_workbook

from rehab_excel.daily_input_journal import JOURNAL_SHEET
from rehab_excel.daily_input_journal_sheet import (
    append_native_operation_revision,
    ensure_daily_input_journal_sheet,
    read_native_journal_rows,
)
from rehab_excel.daily_input_operation import (
    DailyInputOperation,
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_journal_vba import (
    JOURNAL_MODULE_CODE,
    JOURNAL_MODULE_NAME,
    install_daily_input_journal,
)
from rehab_excel.vba_form_designer_audit import (
    audit_userform_designers,
    assert_userform_designers_preserved,
)
from rehab_excel.vba_form_storage import (
    form_storage_evidence,
)
from rehab_excel.vba_semantic_audit import FORM_TYPE, audit_vba_components
from tools.prepare_native_probe_copy import create_probe_copy
from tools.prepare_daily_input_journal_prepared import (
    _source_date,
    prepare_prepared_handoff,
)
from tools.diagnose_excel_journal_vs_noedit_ole import (
    assert_matching_source_relative_change_footprints,
    compare_form_storage_evidence,
    stream_change_categories,
)


class PreparedVbaHookDiagnosticError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _configure_excel(app) -> None:
    app.AutomationSecurity = 3
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.ScreenUpdating = False
    app.AskToUpdateLinks = False


def _excel_factory():
    if sys.platform != "win32":
        raise PreparedVbaHookDiagnosticError("Windows + native Excel required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        raise PreparedVbaHookDiagnosticError("pywin32 is required") from None
    return win32com.client.DispatchEx("Excel.Application")


def _service_date(source: Path) -> str:
    workbook = None
    try:
        workbook = load_workbook(source, read_only=True, data_only=False, keep_vba=True)
        return _source_date(workbook["DAILY_INPUT"]["B2"].value)
    except Exception:
        raise PreparedVbaHookDiagnosticError("DAILY_INPUT date could not be read") from None
    finally:
        if workbook is not None:
            workbook.close()


def _fictional_operations(service_date: str):
    for operation_id in ("FICTIONAL-PREPARED-HOOK-A", "FICTIONAL-PREPARED-HOOK-B"):
        operation = DailyInputOperation(
            operation_id=operation_id,
            source_sha256="a" * 64,
            service_date=service_date,
        )
        yield operation
        operation = advance_daily_input_operation(
            operation, Phase.FAILED, failure_code="FICTIONAL_DIAGNOSTIC",
        )
        yield operation
        operation = advance_daily_input_operation(
            operation, Phase.ROLLED_BACK,
            rollback_verified=True,
        )
        yield operation


def _canonical_vba_source(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")


def _assert_no_auto_open_macros(workbook) -> None:
    components = workbook.VBProject.VBComponents
    for index in range(1, int(components.Count) + 1):
        component = components(index)
        module = component.CodeModule
        count = int(module.CountOfLines)
        if count <= 0:
            continue
        source = str(module.Lines(1, count))
        if re.search(
            r"(?im)^\s*(?:(?:Public|Private|Friend)\s+)?Sub\s+Auto_Open\s*(?:\(|$)",
            source,
        ):
            raise PreparedVbaHookDiagnosticError(
                "Refusing to enable macros while an Auto_Open procedure exists"
            )


def _module_source(workbook) -> str:
    component = workbook.VBProject.VBComponents(JOURNAL_MODULE_NAME)
    module = component.CodeModule
    count = int(module.CountOfLines)
    source = str(module.Lines(1, count)) if count else ""
    return _canonical_vba_source(source)


def _provision_copy(path: Path, *, excel_factory, expected_forms: tuple[str, ...]) -> tuple:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly) or bool(workbook.VBASigned):
            raise PreparedVbaHookDiagnosticError("Disposable XLSM is not writable or is signed")
        _assert_no_auto_open_macros(workbook)
        components = audit_vba_components(workbook)
        forms = tuple(item.name for item in components if item.component_type == FORM_TYPE)
        if tuple(sorted(x.casefold() for x in forms)) != tuple(
            sorted(x.casefold() for x in expected_forms)
        ):
            raise PreparedVbaHookDiagnosticError("UserForm inventory changed before trial")
        designers = audit_userform_designers(workbook, expected_form_names=forms)
        if any(
            str(workbook.Worksheets(index).Name) == JOURNAL_SHEET
            for index in range(1, int(workbook.Worksheets.Count) + 1)
        ):
            raise PreparedVbaHookDiagnosticError("Source already contains a journal")
        journal = ensure_daily_input_journal_sheet(workbook)
        service_date = _service_date(path)
        for operation in _fictional_operations(service_date):
            if not append_native_operation_revision(workbook, operation):
                raise PreparedVbaHookDiagnosticError("Fictional history append was not new")
        install_daily_input_journal(workbook.VBProject)
        workbook.Save()
        if _module_source(workbook) != _canonical_vba_source(JOURNAL_MODULE_CODE):
            raise PreparedVbaHookDiagnosticError("VBA journal module source readback mismatch")
        rows = read_native_journal_rows(journal)
        if len(rows) != 6:
            raise PreparedVbaHookDiagnosticError("Six fictional terminal history rows not staged")
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None
        return forms, designers, rows
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def _second_control_save(path: Path, *, excel_factory) -> None:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly):
            raise PreparedVbaHookDiagnosticError("Control copy unexpectedly read-only")
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def _make_handoff(path: Path, folder: Path) -> tuple[dict[str, object], Path, Path, Path]:
    payload = folder / "fictional-prepared-payload.tsv"
    receipt_path = folder / "fictional-prepared-response.json"
    handoff_path = folder / "fictional-prepared-handoff.json"
    row_path = folder / "fictional-prepared-row.tsv"
    payload.write_text("PATCH\tFICTIONAL_ONLY\n", encoding="utf-8")
    service_date = _service_date(path)
    receipt_path.write_text(json.dumps({
        "ok": True,
        "source_unchanged": True,
        "date": service_date,
        "source_sha256_before": _sha256(path),
        "payload_path": str(payload.resolve()),
    }, ensure_ascii=True, sort_keys=True), encoding="utf-8")
    handoff = prepare_prepared_handoff(
        path, receipt_path, payload,
        operation_id="FICTIONAL-PREPARED-VBA-" + _sha256(path)[:16],
    )
    handoff_path.write_text(
        json.dumps(handoff, ensure_ascii=True, sort_keys=True, separators=(",", ":")),
        encoding="utf-8",
    )
    row = handoff["prepared_row"]
    if len(row) != 5:
        raise PreparedVbaHookDiagnosticError("Canonical prepared row shape invalid")
    row_path.write_text("\t".join(str(value) for value in row) + "\n", encoding="utf-8")
    return handoff, payload, handoff_path, row_path


def _run_vba_hook(
    path: Path,
    *,
    excel_factory,
    handoff: dict[str, object],
    payload: Path,
    handoff_path: Path,
    row_path: Path,
) -> tuple:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        # Auto_Open was excluded in the macro-disabled provisioning pass.
        # Workbook events remain disabled; only the named probe is run.
        app.AutomationSecurity = 1
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly):
            raise PreparedVbaHookDiagnosticError("VBA hook copy unexpectedly read-only")
        preflight_script = REPO_ROOT / "Python" / "tools" / "prepare_daily_input_journal_prepared.py"
        readback_script = REPO_ROOT / "Python" / "tools" / "verify_daily_input_journal_prepared.py"
        macro = f"'{workbook.Name}'!RunPreparedDailyInputJournalProbe"
        result = app.Run(
            macro,
            str(row_path.resolve()),
            str(handoff_path.resolve()),
            str(payload.resolve()),
            str(preflight_script.resolve()),
            str(readback_script.resolve()),
        )
        if not bool(result):
            raise PreparedVbaHookDiagnosticError("VBA PREPARED journal hook returned False")
        if bool(app.Run(f"'{workbook.Name}'!DailyInputJournalApplyEnabled")):
            raise PreparedVbaHookDiagnosticError("Operational Apply gate was unexpectedly enabled")
        journal = workbook.Worksheets(JOURNAL_SHEET)
        rows = read_native_journal_rows(journal)
        if len(rows) != int(handoff["existing_journal_rows"]) + 1:
            raise PreparedVbaHookDiagnosticError("VBA PREPARED row count mismatch")
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None
        return rows
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def _readonly_verify(
    path: Path,
    *,
    excel_factory,
    expected_forms: tuple[str, ...],
    baseline_designers: tuple,
    expected_rows: tuple[tuple, ...],
) -> tuple:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise PreparedVbaHookDiagnosticError("Final readback was not read-only")
        saved_designers = audit_userform_designers(
            workbook, expected_form_names=expected_forms,
        )
        assert_userform_designers_preserved(baseline_designers, saved_designers)
        if _module_source(workbook) != JOURNAL_MODULE_CODE:
            raise PreparedVbaHookDiagnosticError("Journal VBA source changed after reopen")
        journal = workbook.Worksheets(JOURNAL_SHEET)
        if int(journal.Visible) != 2 or not bool(journal.ProtectContents):
            raise PreparedVbaHookDiagnosticError("Reopened journal is not protected/veryHidden")
        rows = read_native_journal_rows(journal)
        if rows != expected_rows:
            raise PreparedVbaHookDiagnosticError("Final native journal history mismatch")
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None
        return saved_designers
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def diagnose_prepared_vba_hook(source_path: str | Path, *, excel_factory=None) -> dict[str, object]:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise PreparedVbaHookDiagnosticError("Existing demo XLSM source required")
    if excel_factory is None:
        excel_factory = _excel_factory
    source_before = _sha256(source)

    # Inspect source UserForms once; every save is compared to this baseline.
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        workbook = app.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise PreparedVbaHookDiagnosticError("Source inspection was not read-only")
        source_components = audit_vba_components(workbook)
        forms = tuple(item.name for item in source_components if item.component_type == FORM_TYPE)
        if not forms:
            raise PreparedVbaHookDiagnosticError("No UserForms available for independent comparison")
        baseline_designers = audit_userform_designers(workbook, expected_form_names=forms)
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    with TemporaryDirectory(prefix="diomil_prepared_vba_", dir=source.parent) as temp:
        folder = Path(temp)
        control = folder / "control.xlsm"
        trial = folder / "vba-prepared.xlsm"
        create_probe_copy(source, control)
        create_probe_copy(source, trial)

        control_forms, control_baseline, control_rows = _provision_copy(
            control, excel_factory=excel_factory, expected_forms=forms,
        )
        trial_forms, trial_baseline, trial_rows = _provision_copy(
            trial, excel_factory=excel_factory, expected_forms=forms,
        )
        if control_forms != trial_forms or control_rows != trial_rows:
            raise PreparedVbaHookDiagnosticError("Control and trial baselines differ")
        if len(control_rows) != 6:
            raise PreparedVbaHookDiagnosticError("Fictional terminal history baseline incomplete")

        # Match the trial's second native save with an untouched control save.
        _second_control_save(control, excel_factory=excel_factory)
        handoff, payload, handoff_path, row_path = _make_handoff(trial, folder)
        trial_rows = _run_vba_hook(
            trial, excel_factory=excel_factory, handoff=handoff,
            payload=payload, handoff_path=handoff_path, row_path=row_path,
        )
        if trial_rows[:-1] != control_rows:
            raise PreparedVbaHookDiagnosticError("Prior journal history changed during VBA append")
        if tuple(trial_rows[-1]) != tuple(handoff["prepared_row"]):
            raise PreparedVbaHookDiagnosticError("VBA PREPARED row differs from canonical handoff")

        control_designers = _readonly_verify(
            control, excel_factory=excel_factory, expected_forms=forms,
            baseline_designers=control_baseline, expected_rows=control_rows,
        )
        trial_designers = _readonly_verify(
            trial, excel_factory=excel_factory, expected_forms=forms,
            baseline_designers=trial_baseline, expected_rows=trial_rows,
        )
        assert_userform_designers_preserved(control_designers, trial_designers)

        source_ole = form_storage_evidence(source, forms)
        control_ole = form_storage_evidence(control, forms)
        trial_ole = form_storage_evidence(trial, forms)
        changes = compare_form_storage_evidence(control_ole, trial_ole)
        categories = stream_change_categories(control_ole, trial_ole)
        changed_streams = assert_matching_source_relative_change_footprints(
            source_ole, control_ole, trial_ole,
        )

    if _sha256(source) != source_before:
        raise PreparedVbaHookDiagnosticError("Original source changed during VBA hook diagnostic")
    return {
        "form_count": len(forms),
        "designer_forms_verified": len(trial_designers),
        "differences": changes,
        "categories": categories,
        "source_relative_footprint_streams": changed_streams,
        "fictional_revisions_verified": len(control_rows),
        "prepared_vba_rows": len(trial_rows) - len(control_rows),
        "prepared_journal_rows": len(trial_rows),
        "source_unchanged": True,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Execute only the VBA PREPARED journal hook on disposable synthetic copies."
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Native VBA hook diagnostic requires Windows + Excel")
        return 2
    try:
        result = diagnose_prepared_vba_hook(args.source)
    except Exception as exc:
        detail = " ".join(str(exc).replace(str(args.source.resolve()), "<SOURCE>").split())[:500]
        print("SAFETY STOP: DAILY_INPUT PREPARED VBA HOOK FAILED")
        print("DIAGNOSTIC: " + detail)
        return 2
    print("DAILY_INPUT PREPARED VBA HOOK DIAGNOSTIC GREEN")
    print(f"UserForms compared: {result['form_count']}")
    print(f"Designer structural readback: {result['designer_forms_verified']} forms matched")
    for name, count in result["differences"].items():
        print(f"{name}={count}")
    for name, count in result["categories"].items():
        print(f"{name}={count}")
    print("Source-relative OLE change footprint matched: YES")
    print(f"Source-relative changed stream identities matched: {result['source_relative_footprint_streams']}")
    print(f"Fictional journal revisions saved/reopened: {result['fictional_revisions_verified']}")
    print(f"Fictional PREPARED VBA hook saved/reopened: {result['prepared_vba_rows']}")
    print(f"Total journal revisions after readback: {result['prepared_journal_rows']}")
    print("VBA hook executed: YES")
    print("Native Excel read-only verifier: GREEN")
    print("Operational Apply/Restore executed: NO")
    print("Source untouched: YES")
    print("Both temporary copies and synthetic sidecars removed")
    print("DIAGNOSTIC ONLY: this is not Apply/Restore or clinical approval")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
