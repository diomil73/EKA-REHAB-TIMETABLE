"""Read-only native Excel inventory for Gate 2. NEVER invokes VBA actions.

Designed for the known isolated fictional demo. Existing action modules may be
absent; report their absence rather than claiming a native guard pass.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT / "Python") not in sys.path:
    sys.path.insert(0, str(ROOT / "Python"))

from rehab_excel.authoritative_commit import file_sha256
from tools.prepare_native_probe_copy import create_probe_copy
from tools.diagnose_daily_input_full_rollback_native import _source_forms
from tools.diagnose_daily_input_prepared_vba_hook import _provision_copy
from rehab_excel.daily_input_journal_vba import (
    JOURNAL_MODULE_NAME,
    JOURNAL_MODULE_CODE,
    canonicalize_vba_source,
)
from tools.diagnose_daily_input_prepared_vba_hook import (
    _configure_excel, _excel_factory,
)

ACTION_MODULES = (
    "modDailyInputAction",
    "modDailyChangeLedger",
    "modGlobalReplacementAction",
    "modDoubleBookingReview",
    "modCloseDayAction",
    "modDayAmendmentAction",
)
DISABLED_GATE = re.compile(
    r"(?im)^\s*Public\s+Function\s+DailyInputJournalApplyEnabled\s*\(\s*\)\s*As\s+Boolean\s*$"
)
FALSE_ASSIGNMENT = re.compile(
    r"(?im)^\s*DailyInputJournalApplyEnabled\s*=\s*False\s*$"
)


class NativeGuardInventoryError(RuntimeError):
    pass


def _read_module(components, name: str) -> str | None:
    try:
        component = components(name)
    except Exception:
        return None
    code = component.CodeModule
    count = int(code.CountOfLines)
    return str(code.Lines(1, count)) if count else ""


def inspect_guard_source(source: str) -> bool:
    """Check the precise disabled predicate, not an unrelated textual FALSE."""
    matching = list(DISABLED_GATE.finditer(source))
    if len(matching) != 1:
        return False
    tail = source[matching[0].end():]
    end = re.search(r"(?im)^\s*End\s+Function\s*$", tail)
    if end is None:
        return False
    body = tail[:end.start()]
    statements = [
        line.strip()
        for line in body.splitlines()
        if line.strip() and not line.lstrip().startswith("'")
    ]
    return statements == ["DailyInputJournalApplyEnabled = False"]


def inspect_native(source_path: str | Path, *, excel_factory=_excel_factory):
    source = Path(source_path).resolve()
    if not source.is_file() or source.suffix.casefold() != ".xlsm":
        raise NativeGuardInventoryError("Synthetic XLSM required")
    before = file_sha256(source)
    forms = _source_forms(source, excel_factory=excel_factory)
    with TemporaryDirectory(prefix="diomil_guard_inventory_", dir=source.parent) as tmp:
        work = Path(tmp) / "guard-inventory.xlsm"
        create_probe_copy(source, work)
        _provision_copy(work, excel_factory=excel_factory, expected_forms=forms)
        app = None
        book = None
        try:
            app = excel_factory()
            _configure_excel(app)
            book = app.Workbooks.Open(
                str(work), UpdateLinks=0, ReadOnly=True, AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
            if not bool(book.ReadOnly):
                raise NativeGuardInventoryError("Read-only inspection not confirmed")
            components = book.VBProject.VBComponents
            guard = _read_module(components, JOURNAL_MODULE_NAME)
            if guard is None or not inspect_guard_source(guard):
                raise NativeGuardInventoryError("Installed journal gate is not strictly FALSE")
            if canonicalize_vba_source(guard) != canonicalize_vba_source(JOURNAL_MODULE_CODE):
                raise NativeGuardInventoryError("Installed journal module differs from source")
            present = tuple(
                name for name in ACTION_MODULES
                if _read_module(components, name) is not None
            )
        finally:
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
    if file_sha256(source) != before:
        raise NativeGuardInventoryError("Source changed during inspection")
    return present


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Synthetic Gate 2 native read-only guard inventory")
    parser.add_argument("--source", required=True, type=Path)
    args = parser.parse_args(argv)
    if sys.platform != "win32":
        print("SAFETY STOP: Windows Excel required")
        return 2
    try:
        present = inspect_native(args.source)
    except Exception:
        print("SAFETY STOP: GATE 2 NATIVE GUARD INVENTORY FAILED")
        return 2
    print("GATE 2 NATIVE READ-ONLY JOURNAL GUARD GREEN")
    print("Installed journal source: VERIFIED; Apply gate: FALSE")
    print(f"Action modules present: {len(present)}/{len(ACTION_MODULES)}")
    print("Action modules missing: " + str(len(ACTION_MODULES) - len(present)))
    print("NO VBA MACROS EXECUTED; ONLY DISPOSABLE COPY SAVED; SOURCE UNCHANGED")
    print("INVENTORY ONLY: DOES NOT PROVE NEGATIVE ACTION EXECUTION")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
