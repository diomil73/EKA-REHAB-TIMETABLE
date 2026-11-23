"""Run a full-workbook rollback candidate/reopen rehearsal on synthetic copies.

This diagnostic reuses the previously native-proven VBA PREPARED hook. It then
introduces a FICTIONAL sheet change and appends FAILED via the existing Python
COM adapter, assembles a separate rollback candidate from the PREPARED snapshot,
verifies it using strict package comparison and fresh macro-disabled Excel, and
adds a ROLLED_BACK marker to a SECOND synthetic candidate for final readback.

Only the PREPARED step is VBA-owned. FAILED is appended by Python COM, and
ROLLED_BACK is serialized in a disposable ZIP fixture, NOT by operational VBA.
This NEVER executes the Apply or Restore macro, never promotes a candidate,
never modifies the supplied source, and is NOT a clinical approval.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_full_rollback_audit import (  # noqa: E402
    _package_members,
    _read_journal,
    audit_full_workbook_rollback,
    audit_restored_business_state_before_rollback,
)
from rehab_excel.daily_input_rollback_candidate import (  # noqa: E402
    _journal_xml_for_history,
    build_rollback_candidate,
)
from rehab_excel.daily_input_journal import (  # noqa: E402
    JOURNAL_SHEET,
    append_operation_journal,
    decode_operation_journal,
)
from rehab_excel.daily_input_journal_sheet import (  # noqa: E402
    append_native_operation_revision,
    read_native_journal_rows,
)
from rehab_excel.daily_input_operation import (  # noqa: E402
    DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.vba_semantic_audit import FORM_TYPE, audit_vba_components  # noqa: E402
from rehab_excel.vba_form_designer_audit import (  # noqa: E402
    audit_userform_designers,
)
from tools.prepare_native_probe_copy import create_probe_copy  # noqa: E402
from tools.verify_daily_input_journal_prepared import (  # noqa: E402
    verify_native_excel_readonly,
)
from tools.diagnose_daily_input_prepared_vba_hook import (  # noqa: E402
    PreparedVbaHookDiagnosticError,
    _configure_excel,
    _excel_factory,
    _make_handoff,
    _provision_copy,
    _readonly_verify,
    _run_vba_hook,
)


class NativeRollbackDiagnosticError(RuntimeError):
    """A synthetic native Excel rollback gate failed closed."""


def _source_forms(path: Path, *, excel_factory) -> tuple[str, ...]:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=True, AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise NativeRollbackDiagnosticError("Source inspection not read-only")
        components = audit_vba_components(workbook)
        forms = tuple(item.name for item in components if item.component_type == FORM_TYPE)
        if len(forms) != 10:
            raise NativeRollbackDiagnosticError("Expected exactly 10 source UserForms")
        audit_userform_designers(workbook, expected_form_names=forms)
        return forms
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


def _write_fictional_failed(
    path: Path,
    *,
    excel_factory,
    prepared_rows: tuple[tuple, ...],
) -> tuple[tuple, ...]:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)  # Macros disabled, events disabled throughout.
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=False, AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly) or bool(workbook.VBASigned):
            raise NativeRollbackDiagnosticError("Disposable failure fixture not writable")
        journal = workbook.Worksheets(JOURNAL_SHEET)
        if read_native_journal_rows(journal) != prepared_rows:
            raise NativeRollbackDiagnosticError("PREPARED history changed before synthetic fault")
        current = decode_operation_journal(prepared_rows)[-1]
        if current.phase != Phase.PREPARED:
            raise NativeRollbackDiagnosticError("Synthetic fault must begin at PREPARED")

        # Isolated, unmistakably fictional mutation. No patient or treatment
        # field is addressed, and only a disposable copy is writable.
        fault = workbook.Worksheets.Add(
            After=workbook.Worksheets(workbook.Worksheets.Count)
        )
        fault.Name = "DIOMIL_SYNTHETIC_FAULT"
        fault.Cells(1, 1).Value2 = "FICTIONAL_CHANGE_ONLY"
        next_state = advance_daily_input_operation(
            current, Phase.FAILED, failure_code="FICTIONAL_NATIVE_FAILURE"
        )
        if not append_native_operation_revision(workbook, next_state):
            raise NativeRollbackDiagnosticError("Synthetic FAILED revision was not new")
        workbook.Save()
        rows = read_native_journal_rows(journal)
        if rows[:-1] != prepared_rows or rows[-1][2] != Phase.FAILED:
            raise NativeRollbackDiagnosticError("FAILED history stage does not match")
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


def _finalize_synthetic_rollback(
    prefinal_path: Path,
    final_path: Path,
) -> tuple[tuple, ...]:
    """Technical-only ZIP fixture. Not an operational VBA journal writer."""
    if prefinal_path.resolve() == final_path.resolve() or final_path.exists():
        raise NativeRollbackDiagnosticError("Terminal fixture output path unavailable")
    rows = _read_journal(prefinal_path)
    states = decode_operation_journal(rows)
    if not states or states[-1].phase != Phase.FAILED:
        raise NativeRollbackDiagnosticError("Synthetic terminal fixture missing FAILED")
    rolled_back = advance_daily_input_operation(
        states[-1], Phase.ROLLED_BACK, rollback_verified=True,
    )
    expected = append_operation_journal(rows, rolled_back)
    if expected[:-1] != rows or expected[-1][2] != Phase.ROLLED_BACK:
        raise NativeRollbackDiagnosticError("Synthetic terminal revision rejected")

    parts, journal_part = _package_members(prefinal_path, require_vba=True)
    journal_xml = _journal_xml_for_history(parts[journal_part], expected)
    created = False
    try:
        with ZipFile(prefinal_path, "r") as source_archive:
            with ZipFile(final_path, "x") as target_archive:
                created = True
                for info in source_archive.infolist():
                    data = journal_xml if info.filename == journal_part else parts[info.filename]
                    target_archive.writestr(info, data)
        if _read_journal(final_path) != expected:
            raise NativeRollbackDiagnosticError("Synthetic terminal journal failed local readback")
    except Exception:
        if created:
            try:
                final_path.unlink()
            except OSError:
                pass
        raise
    return expected


def diagnose_native_rollback(
    source_path: str | Path,
    *,
    excel_factory=None,
) -> dict[str, int | bool]:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeRollbackDiagnosticError("Existing synthetic demo XLSM required")
    if excel_factory is None:
        excel_factory = _excel_factory
    source_sha = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(prefix="diomil_rollback_native_", dir=source.parent) as folder_name:
        folder = Path(folder_name)
        working = folder / "vba-prepared.xlsm"
        create_probe_copy(source, working)
        working_forms, baseline_designers, original_rows = _provision_copy(
            working, excel_factory=excel_factory, expected_forms=forms,
        )
        if working_forms != forms or len(original_rows) != 6:
            raise NativeRollbackDiagnosticError("Fictional native history baseline invalid")
        handoff, payload, handoff_path, row_path = _make_handoff(working, folder)
        prepared_rows = _run_vba_hook(
            working, excel_factory=excel_factory, handoff=handoff,
            payload=payload, handoff_path=handoff_path, row_path=row_path,
        )
        if len(prepared_rows) != 7 or prepared_rows[-1][2] != Phase.PREPARED:
            raise NativeRollbackDiagnosticError("VBA PREPARED append was not verified")

        snapshot = folder / "prepared-snapshot.xlsm"
        failed = folder / "failed-fixture.xlsm"
        candidate = folder / "restored-candidate.xlsm"
        terminal = folder / "terminal-fictional.xlsm"
        create_probe_copy(working, snapshot)
        create_probe_copy(working, failed)
        snapshot_sha = file_sha256(snapshot)
        failed_rows = _write_fictional_failed(
            failed, excel_factory=excel_factory, prepared_rows=prepared_rows,
        )
        if len(failed_rows) != len(prepared_rows) + 1:
            raise NativeRollbackDiagnosticError("Synthetic FAILED revision missing")

        result = build_rollback_candidate(
            snapshot, failed, candidate,
            expected_snapshot_sha256=snapshot_sha,
        )
        pre = audit_restored_business_state_before_rollback(
            snapshot, candidate, expected_snapshot_sha256=snapshot_sha,
        )
        if pre.restored_journal_revisions != len(failed_rows):
            raise NativeRollbackDiagnosticError("Recovery candidate journal length invalid")
        _readonly_verify(
            candidate, excel_factory=excel_factory, expected_forms=forms,
            baseline_designers=baseline_designers, expected_rows=failed_rows,
        )
        verify_native_excel_readonly(candidate, failed_rows)

        terminal_rows = _finalize_synthetic_rollback(candidate, terminal)
        post = audit_full_workbook_rollback(
            snapshot, terminal, expected_snapshot_sha256=snapshot_sha,
        )
        if post.restored_journal_revisions != len(terminal_rows):
            raise NativeRollbackDiagnosticError("Terminal journal length invalid")
        _readonly_verify(
            terminal, excel_factory=excel_factory, expected_forms=forms,
            baseline_designers=baseline_designers, expected_rows=terminal_rows,
        )
        verify_native_excel_readonly(terminal, terminal_rows)
        if (file_sha256(snapshot) != snapshot_sha
                or result.candidate_sha256 != file_sha256(candidate)):
            raise NativeRollbackDiagnosticError("Synthetic source/candidate changed during checks")

    if file_sha256(source) != source_sha:
        raise NativeRollbackDiagnosticError("Original source was modified")
    return {
        "forms": len(forms),
        "fictional_prior": len(original_rows),
        "vba_prepared": 1,
        "com_failed": 1,
        "zip_synthetic_rolled_back": 1,
        "journal_after": len(terminal_rows),
        "nonjournal_parts": post.checked_nonjournal_parts,
        "source_unchanged": True,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Native Excel rollback candidate diagnostic, disposable synthetic copies ONLY."
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Native rollback diagnostic requires Windows + Excel")
        return 2
    try:
        result = diagnose_native_rollback(args.source)
    except Exception:
        # Do not disclose patient/file paths, workbook data, or private codes.
        print("SAFETY STOP: DAILY_INPUT NATIVE ROLLBACK CANDIDATE FAILED")
        return 2
    print("DAILY_INPUT NATIVE ROLLBACK CANDIDATE DIAGNOSTIC GREEN")
    print("UserForms checked in native Excel: " + str(result["forms"]))
    print("Fictional preceding journal revisions: " + str(result["fictional_prior"]))
    print("PREPARED appended by VBA: YES")
    print("FAILED appended by Python COM: YES")
    print("ROLLED_BACK fictional ZIP marker: YES (NOT VBA)")
    print("Total fictional journal revisions: " + str(result["journal_after"]))
    print("Nonjournal workbook members matched: " + str(result["nonjournal_parts"]))
    print("Candidate native read-only reopen: GREEN")
    print("Terminal native read-only reopen: GREEN")
    print("Operational Apply/Restore executed: NO")
    print("Source untouched: YES")
    print("Temporary files removed: YES")
    print("DIAGNOSTIC ONLY: no clinical approval or operational VBA rollback proof")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
