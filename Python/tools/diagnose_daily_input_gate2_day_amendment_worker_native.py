"""Native negative test of the private closed-day amendment writer via disposable shim.

The private PersistDaySavePayload worker checks the disabled journal predicate
before reading the deliberately nonexistent payload. The public
StartClosedDayAmendment UI wrapper is never invoked.
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
from rehab_excel.day_amendment_action_vba import (
    DAY_AMENDMENT_MODULE_NAME,
    DAY_AMENDMENT_MODULE_CODE,
)
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


class NativeDayAmendmentGuardError(RuntimeError):
    pass


SHIM_NAME = "DIOMIL_TestInvokeDayAmendmentGuard"
SHIM_SOURCE = (
    "Public Function DIOMIL_TestInvokeDayAmendmentGuard(ByVal payloadPath As String) As Boolean\n"
    "    DIOMIL_TestInvokeDayAmendmentGuard = PersistDaySavePayload(payloadPath)\n"
    "End Function"
)


def _provision_day_amendment_module(workbook) -> None:
    project = workbook.VBProject
    try:
        existing = project.VBComponents(DAY_AMENDMENT_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        project.VBComponents.Remove(existing)
    component = project.VBComponents.Add(1)
    component.Name = DAY_AMENDMENT_MODULE_NAME
    component.CodeModule.AddFromString(DAY_AMENDMENT_MODULE_CODE)
    workbook.Save()
    installed = _read_module(workbook, DAY_AMENDMENT_MODULE_NAME)
    if canonicalize_vba_source(installed) != canonicalize_vba_source(DAY_AMENDMENT_MODULE_CODE):
        detail = _module_source_mismatch(DAY_AMENDMENT_MODULE_CODE, installed)
        raise NativeDayAmendmentGuardError(
            "Provisioned day amendment module readback mismatch; " + detail
        )
    count = int(component.CodeModule.CountOfLines)
    component.CodeModule.InsertLines(count + 1, SHIM_SOURCE)
    workbook.Save()
    shim_readback = str(component.CodeModule.Lines(count + 1, 3))
    if SHIM_NAME not in shim_readback or "PersistDaySavePayload" not in shim_readback:
        raise NativeDayAmendmentGuardError("Disposable day amendment shim readback mismatch")


def run_locked_day_amendment_worker(
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
            raise NativeDayAmendmentGuardError(
                "Disposable day amendment copy unexpectedly read-only"
            )
        installed = _read_module(workbook, DAY_AMENDMENT_MODULE_NAME)
        base = installed.split("Public Function " + SHIM_NAME, 1)[0]
        if canonicalize_vba_source(base) != canonicalize_vba_source(DAY_AMENDMENT_MODULE_CODE):
            detail = _module_source_mismatch(DAY_AMENDMENT_MODULE_CODE, base)
            raise NativeDayAmendmentGuardError(
                "Installed day amendment module differs from verified source; " + detail
            )
        if missing_payload.exists():
            raise NativeDayAmendmentGuardError(
                "Negative day amendment payload unexpectedly exists"
            )
        result = app.Run(
            f"'{workbook.Name}'!{SHIM_NAME}",
            str(missing_payload),
        )
        if type(result) is not bool or result is not False:
            raise NativeDayAmendmentGuardError(
                "Day amendment payload worker did not return Boolean False"
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
        raise NativeDayAmendmentGuardError(
            "Disposable workbook package changed during rejected day amendment"
        )
    return True


def diagnose_native_day_amendment_guard(
    source_path: str | Path, *, excel_factory=_excel_factory
) -> None:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeDayAmendmentGuardError("Synthetic XLSM source required")
    source_before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(
        prefix="diomil_gate2_day_amendment_guard_", dir=source.parent
    ) as folder:
        root = Path(folder)
        copy = root / "day-amendment-guard.xlsm"
        missing_payload = root / "MUST_NOT_EXIST.day-amendment"
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
                raise NativeDayAmendmentGuardError(
                    "Disposable provisioning copy is read-only"
                )
            _provision_apply_module(workbook)
            _provision_day_amendment_module(workbook)
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
        run_locked_day_amendment_worker(
            copy, missing_payload, excel_factory=excel_factory
        )
        if missing_payload.exists():
            raise NativeDayAmendmentGuardError(
                "Unexpected day amendment payload creation"
            )
    if file_sha256(source) != source_before:
        raise NativeDayAmendmentGuardError("Original synthetic source changed")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Native locked closed-day amendment worker negative test"
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows native Excel required")
        return 2
    try:
        diagnose_native_day_amendment_guard(args.source)
    except Exception as exc:
        print("SAFETY STOP: GATE 2 NATIVE DAY AMENDMENT GUARD FAILED")
        print(f"ERROR_TYPE={type(exc).__name__}")
        message = str(exc).replace("\r", " ").replace("\n", " ")
        print("ERROR_MESSAGE=" + message[:700])
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            print(f"HRESULT=0x{int(hresult) & 0xFFFFFFFF:08X}")
        return 2
    print("GATE 2 NATIVE DAY AMENDMENT WORKER REJECTION GREEN")
    print("PersistDaySavePayload returned Boolean False: YES")
    print("Nonexistent payload was not read or created: YES")
    print("Disposable workbook package unchanged: YES")
    print("Original source unchanged; temporary files removed: YES")
    print("Only private day amendment history worker via disposable-only shim invoked")
    print("Public StartClosedDayAmendment UI wrapper was NOT invoked")
    print("SINGLE WORKER NEGATIVE TEST ONLY: NOT FULL OPERATIONAL TRANSACTION PROOF")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
