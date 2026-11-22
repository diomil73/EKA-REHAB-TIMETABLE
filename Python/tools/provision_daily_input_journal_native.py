"""Provision DAILY_INPUT_JOURNAL in a disposable XLSM copy using native Excel.

The source workbook is never opened or modified. The output must be a new path.
The copy is saved by Excel, closed, reopened read-only, and independently checked
for VBA preservation, original worksheet visibility, and journal schema/protection.
No journal events are written and no operational hooks are installed.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys
from zipfile import BadZipFile, ZipFile

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
for import_root in (REPO_ROOT, PYTHON_ROOT):
    if str(import_root) not in sys.path:
        sys.path.insert(0, str(import_root))

from Python.tools.prepare_native_probe_copy import create_probe_copy
from rehab_excel.daily_input_journal import JOURNAL_HEADERS, JOURNAL_SHEET
from rehab_excel.daily_input_journal_sheet import (
    ensure_daily_input_journal_sheet,
    read_native_journal_rows,
)
from rehab_excel.vba_semantic_audit import (
    FORM_TYPE,
    audit_vba_components,
    assert_vba_semantics_preserved,
)
from rehab_excel.vba_form_storage import (
    form_storage_evidence,
    assert_form_storage_preserved,
)
from rehab_excel.vba_form_designer_audit import (
    audit_userform_designers, assert_userform_designers_preserved,
)


class NativeJournalProvisionError(RuntimeError):
    pass


FORCE_DISABLE_MACROS = 3
XL_SHEET_VERY_HIDDEN = 2


def _file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _vba_sha256(path: Path) -> str:
    try:
        with ZipFile(path) as archive:
            data = archive.read("xl/vbaProject.bin")
    except (BadZipFile, KeyError) as exc:
        raise NativeJournalProvisionError("XLSM VBA project is missing or invalid") from exc
    return sha256(data).hexdigest()


def _worksheet_snapshot(workbook) -> tuple[tuple[str, int], ...]:
    result = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        sheet = workbook.Worksheets(index)
        result.append((str(sheet.Name), int(sheet.Visible)))
    return tuple(result)


def _verify_native_journal(workbook, expected_original_sheets: tuple[tuple[str, int], ...], expected_rows: tuple[tuple, ...]) -> None:
    observed = _worksheet_snapshot(workbook)
    original_names = {name for name, _ in expected_original_sheets}
    observed_original = tuple(item for item in observed if item[0] in original_names)
    if observed_original != expected_original_sheets:
        raise NativeJournalProvisionError("An original worksheet name/order/visibility changed")
    journal_matches = [item for item in observed if item[0] == JOURNAL_SHEET]
    if len(journal_matches) != 1 or journal_matches[0][1] != XL_SHEET_VERY_HIDDEN:
        raise NativeJournalProvisionError("Journal sheet is missing, duplicated, or not veryHidden")

    ws = workbook.Worksheets(JOURNAL_SHEET)
    if not bool(ws.ProtectContents):
        raise NativeJournalProvisionError("Journal worksheet is not protected")
    if tuple(ws.Cells(1, col).Value2 for col in range(1, 6)) != JOURNAL_HEADERS:
        raise NativeJournalProvisionError("Journal headers failed native reopen readback")
    if read_native_journal_rows(ws) != expected_rows:
        raise NativeJournalProvisionError("Journal history changed during native save/reopen")


def provision_daily_input_journal(
    source_path: Path,
    output_path: Path,
    *,
    excel_factory=None,
) -> dict[str, object]:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output or source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise NativeJournalProvisionError("Distinct source and output .xlsm paths are required")
    if not source.is_file():
        raise NativeJournalProvisionError("Source workbook does not exist")
    if output.exists():
        raise NativeJournalProvisionError("Output already exists; refusing overwrite")
    if excel_factory is None:
        if sys.platform != "win32":
            raise NativeJournalProvisionError("Native journal provisioning requires Windows + Excel")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise NativeJournalProvisionError("pywin32 is required") from exc
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    source_hash = _file_sha256(source)
    source_vba_hash = _vba_sha256(source)
    created_copy = False
    app = None
    workbook = None
    stage = "create isolated copy"
    try:
        create_probe_copy(source, output)
        created_copy = True
        stage = "verify source unchanged after copy"
        if _file_sha256(source) != source_hash:
            raise NativeJournalProvisionError("Source changed during isolated-copy creation")

        stage = "start isolated Excel instance"
        app = excel_factory()
        stage = "configure isolated Excel instance"
        # Disable macros before either writable or verification open.
        app.AutomationSecurity = FORCE_DISABLE_MACROS
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False

        stage = "open disposable workbook for provisioning"
        workbook = app.Workbooks.Open(
            str(output), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly):
            raise NativeJournalProvisionError("Excel opened disposable copy read-only")
        stage = "verify unsigned source VBA project"
        if bool(workbook.VBASigned):
            raise NativeJournalProvisionError("Signed VBA project: sheet insertion would invalidate signature")
        stage = "inspect original worksheet roster"
        baseline_sheets = _worksheet_snapshot(workbook)
        stage = "audit baseline VBA component exports"
        baseline_vba = audit_vba_components(workbook)
        form_names = tuple(item.name for item in baseline_vba
                           if item.component_type == FORM_TYPE)
        stage = "audit original embedded UserForm OLE streams"
        original_form_storages = (
            form_storage_evidence(source, form_names) if form_names else ()
        )
        stage = "audit baseline UserForm designer properties"
        original_form_designers = (
            audit_userform_designers(workbook, expected_form_names=form_names)
            if form_names else ()
        )
        if any(name == JOURNAL_SHEET for name, _ in baseline_sheets):
            # Existing journal is validated by ensure(); do not replace or clear it.
            pass
        stage = "create or validate protected journal sheet"
        journal_sheet = ensure_daily_input_journal_sheet(workbook)
        journal_codename = str(journal_sheet.CodeName)
        new_journal = not any(name == JOURNAL_SHEET for name, _ in baseline_sheets)
        expected_journal_rows = read_native_journal_rows(journal_sheet)
        stage = "save disposable workbook in native Excel"
        workbook.Save()
        stage = "close provisioned workbook"
        workbook.Close(SaveChanges=False)
        workbook = None

        stage = "inspect VBA binary hash after native save"
        post_save_vba_hash = _vba_sha256(output)

        # Independent native reopen, read-only. Reopen does not execute macros.
        stage = "reopen disposable workbook read-only"
        workbook = app.Workbooks.Open(
            str(output), UpdateLinks=0, ReadOnly=True,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise NativeJournalProvisionError("Verification reopen was not read-only")
        stage = "verify native journal and original worksheet readback"
        _verify_native_journal(workbook, baseline_sheets, expected_journal_rows)
        stage = "audit saved VBA modules and resources"
        post_save_vba = audit_vba_components(workbook)
        stage = "verify saved UserForm designer control semantics"
        if form_names:
            saved_form_designers = audit_userform_designers(
                workbook, expected_form_names=form_names,
            )
            assert_userform_designers_preserved(
                original_form_designers, saved_form_designers,
            )
        stage = "verify saved embedded UserForm OLE streams"
        saved_form_storages = (
            form_storage_evidence(output, form_names) if form_names else ()
        )
        if form_names:
            assert_form_storage_preserved(original_form_storages, saved_form_storages)
        stage = "compare VBA modules after independent form resource verification"
        protected_modules = assert_vba_semantics_preserved(
            baseline_vba, post_save_vba,
            journal_codename=journal_codename,
            expect_new_document=new_journal,
            independently_verified_form_storages=bool(form_names),
        )
        stage = "close verification workbook and Excel"
        workbook.Close(SaveChanges=False)
        workbook = None
        app.Quit()
        app = None

        stage = "verify final source and VBA hashes"
        if _file_sha256(source) != source_hash:
            raise NativeJournalProvisionError("Source workbook changed during native provisioning")
        if _vba_sha256(output) != post_save_vba_hash:
            raise NativeJournalProvisionError("VBA project changed after verified readback")
        return {
            "copy_created": True,
            "source_unchanged": True,
            "vba_preserved": True,
            "vba_original_components_verified": protected_modules,
            "vba_form_storages_verified": len(form_names),
            "vba_binary_identical": post_save_vba_hash == source_vba_hash,
            "native_save_reopen_verified": True,
            "journal_veryhidden_protected": True,
            "journal_empty": len(expected_journal_rows) == 0,
            "operational_hooks_installed": False,
        }
    except Exception as exc:
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
        if created_copy:
            output.unlink(missing_ok=True)
        message = str(exc).replace(str(source), "<SOURCE>").replace(str(output), "<OUTPUT>")
        message = " ".join(message.split())[:500] or "(no diagnostic message supplied)"
        raise NativeJournalProvisionError(
            f"Stage={stage}; cause={type(exc).__name__}; detail={message}"
        ) from exc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Provision and verify an empty DAILY_INPUT_JOURNAL in a disposable XLSM copy"
    )
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = provision_daily_input_journal(args.source, args.output)
    except Exception as exc:
        detail = " ".join(str(exc).split())[:700]
        print(f"SAFETY STOP: Native journal provisioning failed ({type(exc).__name__}).")
        print(f"DIAGNOSTIC: {detail}")
        return 2
    print("DAILY_INPUT JOURNAL NATIVE PROVISION OK")
    print("Disposable copy created: YES")
    print("Source unchanged: YES")
    print("Original VBA component code/resources preserved: YES")
    print(f"Verified original VBA components: {result['vba_original_components_verified']}")
    print(f"Original UserForm OLE storages verified: {result['vba_form_storages_verified']}")
    print("VBA binary byte-identical: " + ("YES" if result["vba_binary_identical"] else "NO (metadata changed; semantic audit passed)"))
    print("Excel save/close/reopen readback: GREEN")
    print("Journal veryHidden + protected: YES")
    print("Journal empty: " + ("YES" if result["journal_empty"] else "NO (existing history preserved)"))
    print("Apply/Restore hooks installed: NO")
    print("This is a technical probe copy, not an operational workbook.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
