"""Minimal native, copy-only afternoon read-only panel installation.

Only runs with Microsoft Excel, on a NEW separate XLSM, after explicit
synthetic provenance selection. Does not call the broad patient UI installer,
refresh MASTER, or touch source bytes. Does not admit or migrate patients.
"""
from __future__ import annotations

from hashlib import sha256
from pathlib import Path
import shutil
import sys

from rehab_excel.afternoon_native_preflight import (
    AfternoonNativePreflightError, inspect_afternoon_native_candidate,
)
from rehab_excel.afternoon_patient_readonly_vba import (
    PANEL_LABEL, install_afternoon_readonly_panel, patch_afternoon_panel_text,
    _module_text,
)
from rehab_excel.afternoon_sheet_preview import validate_preview_paths
from rehab_excel.afternoon_slots_source import HEADERS, SHEET_NAME


class AfternoonMinimalInstallError(RuntimeError):
    pass


def _hash(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _position(control, left: float, top: float) -> None:
    control.Left = left
    control.Top = top


def _roster(book) -> tuple[str, ...]:
    return tuple(str(book.Worksheets(i).Name) for i in range(1, int(book.Worksheets.Count) + 1))


def _sheet_snapshot(worksheet) -> tuple:
    """Capture formulas as well as cached values for an existing worksheet."""
    used = worksheet.UsedRange
    return (str(used.Address), used.Value2, used.Formula)


def _baseline(book) -> tuple:
    """Conservative native worksheet invariants, no patient values in logs."""
    master = book.Worksheets("MASTER_SCHEDULE")
    planner = book.Worksheets("PATIENT_PLANNER")
    patients = book.Worksheets("PATIENTS")
    master_used = master.UsedRange
    first_column = int(master_used.Column)
    column_count = int(master_used.Columns.Count)
    master_widths = tuple(
        float(master.Columns(index).ColumnWidth)
        for index in range(first_column, first_column + column_count)
    )
    return (
        _roster(book),
        (first_column, column_count, master_widths),
        _sheet_snapshot(master),
        _sheet_snapshot(planner),
        _sheet_snapshot(patients),
    )


def _canonical_vba_source(code: str) -> str:
    """Ignore only native newline representation and trailing empty lines."""
    return code.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")


def _confirm_panel(book, original_runtime: str, original_form: str) -> None:
    project = book.VBProject
    runtime = project.VBComponents("PatientProgramRuntime")
    form = project.VBComponents("frmPatientScheduling")
    expected_runtime, expected_form = patch_afternoon_panel_text(
        original_runtime, original_form
    )
    if _canonical_vba_source(_module_text(runtime)) != _canonical_vba_source(expected_runtime):
        raise AfternoonMinimalInstallError("Afternoon runtime readback mismatch")
    if _canonical_vba_source(_module_text(form)) != _canonical_vba_source(expected_form):
        raise AfternoonMinimalInstallError("Afternoon form code readback mismatch")
    label = form.Designer.Controls(PANEL_LABEL)
    if str(label.Caption) != "Απογευματινή ΦΘ: προβολή μόνο":
        raise AfternoonMinimalInstallError("Afternoon form label was not retained")
    ws = book.Worksheets(SHEET_NAME)
    if int(ws.Visible) != 2:
        raise AfternoonMinimalInstallError("Afternoon technical sheet must be veryHidden")
    if tuple(str(ws.Cells(1, i).Value) for i in range(1, 6)) != HEADERS:
        raise AfternoonMinimalInstallError("Afternoon schema readback mismatch")
    if any(ws.Cells(2, i).Value is not None for i in range(1, 6)):
        raise AfternoonMinimalInstallError("Afternoon sheet must start empty")


def install_afternoon_minimal_preview(
    source: str | Path,
    output: str | Path,
) -> Path:
    """Only add one PM sheet and one read-only panel to a fresh disposable copy.

    The source must already have the correct existing form/runtime. Never
    derive form code by overwriting a different VBA version.
    """
    if sys.platform != "win32":
        raise AfternoonMinimalInstallError("Microsoft Excel on Windows required")
    src, dst = validate_preview_paths(source, output)
    try:
        inspect_afternoon_native_candidate(src)
    except AfternoonNativePreflightError as exc:
        raise AfternoonMinimalInstallError("Source structural gate failed") from exc
    original_sha = _hash(src)
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise AfternoonMinimalInstallError("pywin32 required") from exc

    excel = None
    book = None
    verifier = None
    readonly = None
    try:
        with src.open("rb") as input_file, dst.open("xb") as output_file:
            shutil.copyfileobj(input_file, output_file)
        if _hash(src) != original_sha:
            raise AfternoonMinimalInstallError("Source changed during copy")

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        # Disable auto-open macro execution while making a technical-only copy.
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        book = excel.Workbooks.Open(str(dst), UpdateLinks=0, ReadOnly=False)
        if bool(book.ReadOnly) or not bool(book.HasVBProject):
            raise AfternoonMinimalInstallError("Writable VBA workbook required")
        if SHEET_NAME in _roster(book):
            raise AfternoonMinimalInstallError("Existing afternoon sheet forbidden")
        prior = _baseline(book)
        # Source code/access failures must NOT be mistaken for a missing label.
        runtime = book.VBProject.VBComponents("PatientProgramRuntime")
        form = book.VBProject.VBComponents("frmPatientScheduling")
        runtime_code, form_code = _module_text(runtime), _module_text(form)
        patch_afternoon_panel_text(runtime_code, form_code)
        designer = form.Designer
        # Enumerating controls differentiates a missing label from an inaccessible
        # designer. A COM failure now stops installation instead of falling through.
        existing_names = tuple(
            str(designer.Controls(i).Name)
            for i in range(0, int(designer.Controls.Count))
        )
        if PANEL_LABEL in existing_names:
            raise AfternoonMinimalInstallError("Already installed afternoon panel")

        ws = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        ws.Name = SHEET_NAME
        for i, header in enumerate(HEADERS, 1):
            ws.Cells(1, i).Value = header
        ws.Visible = 2
        install_afternoon_readonly_panel(book.VBProject, position_control=_position)
        current = _baseline(book)
        if current[0] != prior[0] + (SHEET_NAME,) or current[1:] != prior[1:]:
            raise AfternoonMinimalInstallError("Morning/Master/registry in-memory invariant changed")
        _confirm_panel(book, runtime_code, form_code)
        book.Save()
        book.Close(SaveChanges=False)
        book = None
        excel.Quit()
        excel = None

        verifier = win32com.client.DispatchEx("Excel.Application")
        verifier.Visible = False
        verifier.DisplayAlerts = False
        verifier.EnableEvents = False
        verifier.AutomationSecurity = 3  # readback must never execute macros
        readonly = verifier.Workbooks.Open(str(dst), UpdateLinks=0, ReadOnly=True)
        reopened = _baseline(readonly)
        if reopened[1:] != prior[1:]:
            raise AfternoonMinimalInstallError("Source sheet values or MASTER widths changed after reopen")
        if reopened[0] != prior[0] + (SHEET_NAME,):
            raise AfternoonMinimalInstallError("Unexpected sheet roster after reopen")
        _confirm_panel(readonly, runtime_code, form_code)
        if _hash(src) != original_sha:
            raise AfternoonMinimalInstallError("Source changed after native save")
        return dst
    except Exception as exc:
        # Preserve fresh output for diagnosis, never delete or retry in place.
        if isinstance(exc, AfternoonMinimalInstallError):
            raise
        raise AfternoonMinimalInstallError(
            "Minimal afternoon preview failed; preserve new output for inspection"
        ) from exc
    finally:
        for workbook in (readonly, book):
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
        for application in (verifier, excel):
            if application is not None:
                try:
                    application.Quit()
                except Exception:
                    pass
