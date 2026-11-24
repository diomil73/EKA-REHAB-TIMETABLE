"""Native negative test of private CloseDay history writer through a disposable-only shim.

The actual PersistCloseDayPayload worker checks the disabled journal predicate
before reading its deliberately nonexistent payload. The public
CloseOperationalDay UI wrapper is never invoked.
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
from rehab_excel.close_day_action_vba import CLOSE_DAY_MODULE_NAME, CLOSE_DAY_MODULE_CODE
from rehab_excel.daily_input_journal_vba import canonicalize_vba_source
from tools.diagnose_daily_input_full_rollback_native import _source_forms
from tools.diagnose_daily_input_prepared_vba_hook import (
    _configure_excel,
    _excel_factory,
    _provision_copy,
    _module_source_mismatch,
)
from tools.diagnose_daily_input_gate2_apply_worker_native import (
    _provision_apply_module,
    _read_module,
)
from tools.prepare_native_probe_copy import create_probe_copy


class NativeCloseDayGuardError(RuntimeError):
    pass


SHIM_NAME = "DIOMIL_TestInvokeCloseDayGuard"
SHIM_SOURCE = (
    "Public Function DIOMIL_TestInvokeCloseDayGuard(ByVal payloadPath As String) As Boolean\n"
    "    DIOMIL_TestInvokeCloseDayGuard = PersistCloseDayPayload(payloadPath)\n"
    "End Function"
)


def _provision_close_day_module(workbook) -> None:
    project = workbook.VBProject
    try:
        existing = project.VBComponents(CLOSE_DAY_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        project.VBComponents.Remove(existing)
    component = project.VBComponents.Add(1)
    component.Name = CLOSE_DAY_MODULE_NAME
    component.CodeModule.AddFromString(CLOSE_DAY_MODULE_CODE)
    workbook.Save()
    installed = _read_module(workbook, CLOSE_DAY_MODULE_NAME)
    if canonicalize_vba_source(installed) != canonicalize_vba_source(CLOSE_DAY_MODULE_CODE):
        detail = _module_source_mismatch(CLOSE_DAY_MODULE_CODE, installed)
        raise NativeCloseDayGuardError(
            "Provisioned CloseDay module readback mismatch; " + detail
        )
    count = int(component.CodeModule.CountOfLines)
    component.CodeModule.InsertLines(count + 1, SHIM_SOURCE)
    workbook.Save()
    shim_readback = str(component.CodeModule.Lines(count + 1, 3))
    if SHIM_NAME not in shim_readback or "PersistCloseDayPayload" not in shim_readback:
        raise NativeCloseDayGuardError("Disposable CloseDay test shim readback mismatch")


def run_locked_close_day_worker(
    copy: Path, missing_payload: Path, *, excel_factory=_excel_factory
) -> bool:
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
            raise NativeCloseDayGuardError("Disposable CloseDay copy unexpectedly read-only")
        installed = _read_module(workbook, CLOSE_DAY_MODULE_NAME)
        base = installed.split("Public Function " + SHIM_NAME, 1)[0]
        if canonicalize_vba_source(base) != canonicalize_vba_source(CLOSE_DAY_MODULE_CODE):
            detail = _module_source_mismatch(CLOSE_DAY_MODULE_CODE, base)
            raise NativeCloseDayGuardError(
                "Installed CloseDay module differs from verified source; " + detail
            )
        if missing_payload.exists():
            raise NativeCloseDayGuardError("Negative CloseDay payload unexpectedly exists")
        result = app.Run(
            f"'{workbook.Name}'!{SHIM_NAME}",
            str(missing_payload),
        )
        if type(result) is not bool or result is not False:
            raise NativeCloseDayGuardError(
                "CloseDay payload worker did not return Boolean False"
            )
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
        raise NativeCloseDayGuardError(
            "Disposable workbook package changed during rejected CloseDay"
        )
    return True


def diagnose_native_close_day_guard(
    source_path: str | Path, *, excel_factory=_excel_factory
) -> None:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeCloseDayGuardError("Synthetic XLSM source required")
    source_before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(
        prefix="diomil_gate2_close_day_guard_", dir=source.parent
    ) as folder:
        root = Path(folder)
        copy = root / "close-day-guard.xlsm"
        missing_payload = root / "MUST_NOT_EXIST.close-day"
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
                raise NativeCloseDayGuardError("Disposable provisioning copy is read-only")
            _provision_apply_module(workbook)
            _provision_close_day_module(workbook)
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
        run_locked_close_day_worker(
            copy, missing_payload, excel_factory=excel_factory
        )
        if missing_payload.exists():
            raise NativeCloseDayGuardError("Unexpected CloseDay payload creation")
    if file_sha256(source) != source_before:
        raise NativeCloseDayGuardError("Original synthetic source changed")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Native locked CloseDay worker negative test"
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows native Excel required")
        return 2
    try:
        diagnose_native_close_day_guard(args.source)
    except Exception as exc:
        print("SAFETY STOP: GATE 2 NATIVE CLOSE DAY GUARD FAILED")
        print(f"ERROR_TYPE={type(exc).__name__}")
        message = str(exc).replace("\r", " ").replace("\n", " ")
        print("ERROR_MESSAGE=" + message[:700])
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            print(f"HRESULT=0x{int(hresult) & 0xFFFFFFFF:08X}")
        return 2
    print("GATE 2 NATIVE CLOSE DAY WORKER REJECTION GREEN")
    print("PersistCloseDayPayload returned Boolean False: YES")
    print("Nonexistent payload was not read or created: YES")
    print("Disposable workbook package unchanged: YES")
    print("Original source unchanged; temporary files removed: YES")
    print("Only private CloseDay history worker via disposable-only shim invoked")
    print("Public CloseOperationalDay UI wrapper was NOT invoked")
    print("SINGLE WORKER NEGATIVE TEST ONLY: NOT FULL OPERATIONAL TRANSACTION PROOF")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
