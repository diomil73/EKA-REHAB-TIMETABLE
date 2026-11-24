"""Native negative test of the private Restore payload worker through a disposable-only shim.

The actual private ApplyDailyChangeRestorePayload worker checks the disabled
journal predicate before reading its deliberately nonexistent payload. A
temporary public shim is appended only to the disposable VBA component so
Excel can invoke that private function without calling the modal Restore UI.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT / "Python") not in sys.path:
    sys.path.insert(0, str(ROOT / "Python"))

from rehab_excel.authoritative_commit import file_sha256
from rehab_excel.daily_change_ledger_in_place_vba import MODULE_NAME, VBA_CODE
from rehab_excel.daily_input_journal_vba import canonicalize_vba_source
from tools.diagnose_daily_input_full_rollback_native import _source_forms
from tools.diagnose_daily_input_prepared_vba_hook import (
    _configure_excel,
    _excel_factory,
    _provision_copy,
)
from tools.diagnose_daily_input_gate2_apply_worker_native import (
    _provision_apply_module,
    _read_module,
    _module_source_mismatch,
)
from tools.prepare_native_probe_copy import create_probe_copy


class NativeRestoreGuardError(RuntimeError):
    pass


SHIM_NAME = "DIOMIL_TestInvokeRestoreGuard"
SHIM_SOURCE = (
    "Public Function DIOMIL_TestInvokeRestoreGuard(ByVal payloadPath As String) As Boolean\n"
    "    DIOMIL_TestInvokeRestoreGuard = ApplyDailyChangeRestorePayload(payloadPath)\n"
    "End Function"
)


def _provision_restore_module(workbook) -> None:
    project = workbook.VBProject
    try:
        existing = project.VBComponents(MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        project.VBComponents.Remove(existing)
    component = project.VBComponents.Add(1)
    component.Name = MODULE_NAME
    component.CodeModule.AddFromString(VBA_CODE)
    workbook.Save()
    installed = _read_module(workbook, MODULE_NAME)
    if canonicalize_vba_source(installed) != canonicalize_vba_source(VBA_CODE):
        detail = _module_source_mismatch(VBA_CODE, installed)
        raise NativeRestoreGuardError(
            "Provisioned Restore module readback mismatch; " + detail
        )
    count = int(component.CodeModule.CountOfLines)
    component.CodeModule.InsertLines(count + 1, SHIM_SOURCE)
    workbook.Save()
    if SHIM_NAME not in str(component.CodeModule.Lines(count + 1, 3)):
        raise NativeRestoreGuardError("Disposable Restore test shim readback mismatch")


def run_locked_restore_worker(copy: Path, missing_payload: Path, *, excel_factory=_excel_factory) -> bool:
    app = None
    workbook = None
    before = file_sha256(copy)
    try:
        app = excel_factory()
        _configure_excel(app)
        app.AutomationSecurity = 1
        workbook = app.Workbooks.Open(
            str(copy), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly):
            raise NativeRestoreGuardError("Disposable Restore probe copy unexpectedly read-only")
        installed = _read_module(workbook, MODULE_NAME)
        if canonicalize_vba_source(installed) != canonicalize_vba_source(VBA_CODE):
            # The test shim is appended after the verified canonical module.
            base = installed.split("Public Function " + SHIM_NAME, 1)[0]
            if canonicalize_vba_source(base) != canonicalize_vba_source(VBA_CODE):
                detail = _module_source_mismatch(VBA_CODE, base)
                raise NativeRestoreGuardError(
                    "Installed Restore module differs from verified source; " + detail
                )
        if not missing_payload.exists():
            result = app.Run(
                f"'{workbook.Name}'!{SHIM_NAME}",
                str(missing_payload),
            )
        else:
            raise NativeRestoreGuardError("Negative Restore payload unexpectedly exists")
        if type(result) is not bool or result is not False:
            raise NativeRestoreGuardError("Restore payload worker did not return Boolean False")
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
    if file_sha256(copy) != before:
        raise NativeRestoreGuardError("Disposable workbook package changed during rejected Restore")
    return True


def diagnose_native_restore_guard(source_path: str | Path, *, excel_factory=_excel_factory) -> None:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeRestoreGuardError("Synthetic XLSM source required")
    source_before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(prefix="diomil_gate2_restore_guard_", dir=source.parent) as folder:
        root = Path(folder)
        copy = root / "restore-guard.xlsm"
        missing_payload = root / "MUST_NOT_EXIST.restore"
        create_probe_copy(source, copy)
        _provision_copy(copy, excel_factory=excel_factory, expected_forms=forms)
        app = None
        workbook = None
        try:
            app = excel_factory()
            _configure_excel(app)
            workbook = app.Workbooks.Open(
                str(copy), UpdateLinks=0, ReadOnly=False,
                AddToMru=False, IgnoreReadOnlyRecommended=True,
            )
            if bool(workbook.ReadOnly):
                raise NativeRestoreGuardError("Disposable provisioning copy is read-only")
            _provision_apply_module(workbook)
            _provision_restore_module(workbook)
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
        run_locked_restore_worker(copy, missing_payload, excel_factory=excel_factory)
        if missing_payload.exists():
            raise NativeRestoreGuardError("Unexpected Restore payload creation")
    if file_sha256(source) != source_before:
        raise NativeRestoreGuardError("Original synthetic source changed")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Native locked Restore worker negative test")
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows native Excel required")
        return 2
    try:
        diagnose_native_restore_guard(args.source)
    except Exception as exc:
        print("SAFETY STOP: GATE 2 NATIVE RESTORE GUARD FAILED")
        print(f"ERROR_TYPE={type(exc).__name__}")
        message = str(exc).replace("\\r", " ").replace("\\n", " ")
        print("ERROR_MESSAGE=" + message[:500])
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            print(f"HRESULT=0x{int(hresult) & 0xFFFFFFFF:08X}")
        return 2
    print("GATE 2 NATIVE RESTORE WORKER REJECTION GREEN")
    print("ApplyDailyChangeRestorePayload returned Boolean False: YES")
    print("Nonexistent payload was not read or created: YES")
    print("Disposable workbook package unchanged: YES")
    print("Original source unchanged; temporary files removed: YES")
    print("Only private Restore payload worker via disposable-only shim invoked")
    print("SINGLE WORKER NEGATIVE TEST ONLY: NOT FULL OPERATIONAL TRANSACTION PROOF")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
