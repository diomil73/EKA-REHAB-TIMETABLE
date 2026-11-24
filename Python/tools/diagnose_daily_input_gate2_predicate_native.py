"""Execute only the locked journal predicate in native Excel on a disposable copy.

This is NOT an Apply/Restore negative-action test. The only VBA entrypoint
allowed by this diagnostic is DailyInputJournalApplyEnabled(), which is
side-effect-free and must return exactly False.
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
from rehab_excel.daily_input_journal_vba import JOURNAL_MODULE_CODE, canonicalize_vba_source
from tools.diagnose_daily_input_full_rollback_native import _source_forms
from tools.diagnose_daily_input_prepared_vba_hook import (
    _configure_excel,
    _excel_factory,
    _module_source,
    _provision_copy,
)
from tools.prepare_native_probe_copy import create_probe_copy


class NativeGatePredicateError(RuntimeError):
    pass


def execute_locked_predicate(path: Path, *, excel_factory=_excel_factory) -> bool:
    app = None
    workbook = None
    try:
        app = excel_factory()
        _configure_excel(app)
        # The provisioning pass already rejected Auto_Open. Keep workbook
        # events disabled and enable macros only for this named predicate.
        app.AutomationSecurity = 1
        workbook = app.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=False,
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if bool(workbook.ReadOnly):
            raise NativeGatePredicateError("Disposable predicate copy is read-only")
        actual = _module_source(workbook)
        expected = canonicalize_vba_source(JOURNAL_MODULE_CODE)
        if actual != expected:
            raise NativeGatePredicateError("Installed journal module changed before predicate run")
        # This is the sole permitted Application.Run target in this tool.
        result = app.Run(f"'{workbook.Name}'!DailyInputJournalApplyEnabled")
        if type(result) is not bool:
            raise NativeGatePredicateError("Journal predicate did not return a Boolean")
        return result
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


def diagnose_native_predicate(source_path: str | Path, *, excel_factory=_excel_factory) -> None:
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeGatePredicateError("Synthetic XLSM source required")
    before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(prefix="diomil_gate2_predicate_", dir=source.parent) as folder:
        copy = Path(folder) / "gate2-predicate.xlsm"
        create_probe_copy(source, copy)
        _provision_copy(copy, excel_factory=excel_factory, expected_forms=forms)
        enabled = execute_locked_predicate(copy, excel_factory=excel_factory)
        if enabled is not False:
            raise NativeGatePredicateError("Apply gate returned True; safety stop")
    if file_sha256(source) != before:
        raise NativeGatePredicateError("Original synthetic source changed")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Native read-only execution of locked journal predicate")
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows native Excel required")
        return 2
    try:
        diagnose_native_predicate(args.source)
    except Exception:
        print("SAFETY STOP: GATE 2 NATIVE PREDICATE FAILED")
        return 2
    print("GATE 2 NATIVE JOURNAL PREDICATE GREEN")
    print("DailyInputJournalApplyEnabled returned Boolean False: YES")
    print("Only allowlisted predicate executed; no Apply/Restore/action workers invoked")
    print("Disposable copy only; original source unchanged; temporary files removed")
    print("PREDICATE EXECUTION ONLY: DOES NOT PROVE ACTION ENTRYPOINT REJECTION")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
