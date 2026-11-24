"""Native negative test of the direct Apply worker while journal gate is FALSE.

Runs only on a disposable synthetic XLSM. The supplied payload path must not
exist. The expected Boolean False proves the worker exits at its gate before
attempting to read the payload. No UI wrapper or modal MsgBox is invoked.
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
from rehab_excel.daily_input_action_vba import (
    DAILY_INPUT_MODULE_CODE,
    DAILY_INPUT_MODULE_NAME,
)
from rehab_excel.daily_input_journal_vba import canonicalize_vba_source
from tools.diagnose_daily_input_full_rollback_native import _source_forms
from tools.diagnose_daily_input_prepared_vba_hook import (
    _configure_excel,
    _excel_factory,
    _provision_copy,
)
from tools.prepare_native_probe_copy import create_probe_copy


class NativeApplyGuardError(RuntimeError):
    pass


def _read_module(workbook, name: str) -> str:
    try:
        component = workbook.VBProject.VBComponents(name)
        module = component.CodeModule
        count = int(module.CountOfLines)
        return str(module.Lines(1, count)) if count else ""
    except Exception:
        raise NativeApplyGuardError(f"Required VBA module unavailable: {name}") from None


def run_locked_apply_worker(copy: Path, missing_payload: Path, *, excel_factory=_excel_factory) -> bool:
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
            raise NativeApplyGuardError("Disposable Apply probe copy unexpectedly read-only")
        installed = _read_module(workbook, DAILY_INPUT_MODULE_NAME)
        if canonicalize_vba_source(installed) != canonicalize_vba_source(DAILY_INPUT_MODULE_CODE):
            raise NativeApplyGuardError("Installed Apply module differs from verified source")
        if missing_payload.exists():
            raise NativeApplyGuardError("Negative test payload unexpectedly exists")
        # Only the direct worker is called. It must return False at the gate,
        # before checking/reading this deliberately nonexistent payload.
        result = app.Run(
            f"'{workbook.Name}'!ApplyValidatedDailyInputPayload",
            str(missing_payload),
        )
        if type(result) is not bool or result is not False:
            raise NativeApplyGuardError("Apply worker did not return Boolean False")
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
        raise NativeApplyGuardError("Disposable workbook package changed during rejected Apply")
    return True


def diagnose_native_apply_guard(source_path: str | Path, *, excel_factory=_excel_factory) -> None:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeApplyGuardError("Synthetic XLSM source required")
    source_before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(prefix="diomil_gate2_apply_guard_", dir=source.parent) as folder:
        root = Path(folder)
        copy = root / "apply-guard.xlsm"
        missing_payload = root / "MUST_NOT_EXIST.payload"
        create_probe_copy(source, copy)
        _provision_copy(copy, excel_factory=excel_factory, expected_forms=forms)
        run_locked_apply_worker(copy, missing_payload, excel_factory=excel_factory)
        if missing_payload.exists():
            raise NativeApplyGuardError("Unexpected payload creation")
    if file_sha256(source) != source_before:
        raise NativeApplyGuardError("Original synthetic source changed")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Native locked Apply worker negative test")
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows native Excel required")
        return 2
    try:
        diagnose_native_apply_guard(args.source)
    except Exception as exc:
        print("SAFETY STOP: GATE 2 NATIVE APPLY GUARD FAILED")
        print(f"ERROR_TYPE={type(exc).__name__}")
        message = str(exc).replace("\\r", " ").replace("\\n", " ")
        print("ERROR_MESSAGE=" + message[:500])
        hresult = getattr(exc, "hresult", None)
        if hresult is not None:
            try:
                print(f"HRESULT=0x{int(hresult) & 0xFFFFFFFF:08X}")
            except (TypeError, ValueError):
                print("HRESULT=UNAVAILABLE")
        excepinfo = getattr(exc, "excepinfo", None)
        if excepinfo:
            print("EXCEPINFO_PRESENT=YES")
            for index in (0, 1, 2, 5):
                try:
                    value = excepinfo[index]
                except (IndexError, TypeError):
                    continue
                if value not in (None, ""):
                    print(f"EXCEPINFO_{index}={str(value)[:240]}")
        return 2
    print("GATE 2 NATIVE APPLY WORKER REJECTION GREEN")
    print("ApplyValidatedDailyInputPayload returned Boolean False: YES")
    print("Nonexistent payload was not read or created: YES")
    print("Disposable workbook package unchanged: YES")
    print("Original source unchanged; temporary files removed: YES")
    print("No UI wrapper, Restore, CloseDay or other action worker invoked")
    print("SINGLE WORKER NEGATIVE TEST ONLY: NOT FULL OPERATIONAL TRANSACTION PROOF")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
