"""Native negative test of the private Replacement payload worker through a disposable-only shim.

The actual PersistReplacementPayload worker checks the disabled journal
predicate before reading its deliberately nonexistent payload. The test shim
is appended only to the disposable VBA component. No public UI workflow,
replacement preparation, or operational writer is run.
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
from rehab_excel.global_replacement_action_vba import (
    GLOBAL_REPLACEMENT_MODULE_NAME,
    GLOBAL_REPLACEMENT_MODULE_CODE,
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


class NativeReplacementGuardError(RuntimeError):
    pass


SHIM_NAME = "DIOMIL_TestInvokeReplacementGuard"
SHIM_SOURCE = (
    "Public Function DIOMIL_TestInvokeReplacementGuard(ByVal payloadPath As String) As Boolean\n"
    "    DIOMIL_TestInvokeReplacementGuard = PersistReplacementPayload(payloadPath, False)\n"
    "End Function"
)


def _provision_replacement_module(workbook) -> None:
    project = workbook.VBProject
    try:
        existing = project.VBComponents(GLOBAL_REPLACEMENT_MODULE_NAME)
    except Exception:
        existing = None
    if existing is not None:
        project.VBComponents.Remove(existing)
    component = project.VBComponents.Add(1)
    component.Name = GLOBAL_REPLACEMENT_MODULE_NAME
    component.CodeModule.AddFromString(GLOBAL_REPLACEMENT_MODULE_CODE)
    workbook.Save()
    installed = _read_module(workbook, GLOBAL_REPLACEMENT_MODULE_NAME)
    if canonicalize_vba_source(installed) != canonicalize_vba_source(GLOBAL_REPLACEMENT_MODULE_CODE):
        detail = _module_source_mismatch(GLOBAL_REPLACEMENT_MODULE_CODE, installed)
        raise NativeReplacementGuardError(
            "Provisioned Replacement module readback mismatch; " + detail
        )
    count = int(component.CodeModule.CountOfLines)
    component.CodeModule.InsertLines(count + 1, SHIM_SOURCE)
    workbook.Save()
    shim_readback = str(component.CodeModule.Lines(count + 1, 3))
    if SHIM_NAME not in shim_readback or "PersistReplacementPayload" not in shim_readback:
        raise NativeReplacementGuardError("Disposable Replacement test shim readback mismatch")


def run_locked_replacement_worker(
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
            raise NativeReplacementGuardError("Disposable Replacement copy unexpectedly read-only")
        installed = _read_module(workbook, GLOBAL_REPLACEMENT_MODULE_NAME)
        base = installed.split("Public Function " + SHIM_NAME, 1)[0]
        if canonicalize_vba_source(base) != canonicalize_vba_source(GLOBAL_REPLACEMENT_MODULE_CODE):
            detail = _module_source_mismatch(GLOBAL_REPLACEMENT_MODULE_CODE, base)
            raise NativeReplacementGuardError(
                "Installed Replacement module differs from verified source; " + detail
            )
        if missing_payload.exists():
            raise NativeReplacementGuardError("Negative Replacement payload unexpectedly exists")
        result = app.Run(
            f"'{workbook.Name}'!{SHIM_NAME}",
            str(missing_payload),
        )
        if type(result) is not bool or result is not False:
            raise NativeReplacementGuardError(
                "Replacement payload worker did not return Boolean False"
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
        raise NativeReplacementGuardError(
            "Disposable workbook package changed during rejected Replacement"
        )
    return True


def diagnose_native_replacement_guard(
    source_path: str | Path, *, excel_factory=_excel_factory
) -> None:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeReplacementGuardError("Synthetic XLSM source required")
    source_before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(
        prefix="diomil_gate2_replacement_guard_", dir=source.parent
    ) as folder:
        root = Path(folder)
        copy = root / "replacement-guard.xlsm"
        missing_payload = root / "MUST_NOT_EXIST.replacement"
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
                raise NativeReplacementGuardError("Disposable provisioning copy is read-only")
            _provision_apply_module(workbook)
            _provision_replacement_module(workbook)
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
        run_locked_replacement_worker(
            copy, missing_payload, excel_factory=excel_factory
        )
        if missing_payload.exists():
            raise NativeReplacementGuardError("Unexpected Replacement payload creation")
    if file_sha256(source) != source_before:
        raise NativeReplacementGuardError("Original synthetic source changed")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Native locked Replacement worker negative test"
    )
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows native Excel required")
        return 2
    try:
        diagnose_native_replacement_guard(args.source)
    except Exception as exc:
        print("SAFETY STOP: GATE 2 NATIVE REPLACEMENT GUARD FAILED")
        print(f"ERROR_TYPE={type(exc).__name__}")
        message = str(exc).replace("\r", " ").replace("\n", " ")
        print("ERROR_MESSAGE=" + message[:700])
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            print(f"HRESULT=0x{int(hresult) & 0xFFFFFFFF:08X}")
        return 2
    print("GATE 2 NATIVE REPLACEMENT WORKER REJECTION GREEN")
    print("PersistReplacementPayload returned Boolean False: YES")
    print("Nonexistent payload was not read or created: YES")
    print("Disposable workbook package unchanged: YES")
    print("Original source unchanged; temporary files removed: YES")
    print("Only private Replacement payload worker via disposable-only shim invoked")
    print("SINGLE WORKER NEGATIVE TEST ONLY: NOT FULL OPERATIONAL TRANSACTION PROOF")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
