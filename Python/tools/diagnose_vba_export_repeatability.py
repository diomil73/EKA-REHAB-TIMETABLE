"""Read-only repeatability diagnosis for VBA exports in an isolated XLSM copy.

This does NOT install a journal, save a workbook, or publish VBA content.
The source is never opened in Excel and the disposable byte copy is deleted.
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys
from tempfile import TemporaryDirectory

REPO_ROOT = Path(__file__).resolve().parents[2]
for root in (REPO_ROOT, REPO_ROOT / "Python"):
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

from Python.tools.prepare_native_probe_copy import create_probe_copy
from Python.tools.provision_daily_input_journal_native import _file_sha256
from rehab_excel.vba_semantic_audit import (
    VbaSemanticAuditError,
    assert_vba_semantics_preserved,
    audit_vba_components,
)


class RepeatabilityProbeError(RuntimeError):
    pass


def probe_vba_export_repeatability(source_path: Path, *, excel_factory=None) -> dict[str, object]:
    """Twice export unchanged VBComponents from the same read-only open."""
    source = Path(source_path).resolve()
    if source.suffix.casefold() != ".xlsm" or not source.is_file():
        raise RepeatabilityProbeError("Regular XLSM required")
    if excel_factory is None:
        if sys.platform != "win32":
            raise RepeatabilityProbeError("Windows native Excel required")
        try:
            import win32com.client
        except ImportError as exc:
            raise RepeatabilityProbeError("pywin32 required") from exc
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")
    source_hash = _file_sha256(source)
    stage = "prepare disposable copy"
    app = None
    book = None
    try:
        with TemporaryDirectory(prefix="diomil_vba_repeat_") as folder:
            copy = Path(folder) / "repeatability_probe.xlsm"
            create_probe_copy(source, copy)
            stage = "start Excel"
            app = excel_factory()
            stage = "configure Excel"
            app.AutomationSecurity = 3
            app.Visible = False
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.AskToUpdateLinks = False
            stage = "open disposable copy read-only"
            book = app.Workbooks.Open(
                str(copy), UpdateLinks=0, ReadOnly=True,
                AddToMru=False, IgnoreReadOnlyRecommended=True,
            )
            if not bool(book.ReadOnly):
                raise RepeatabilityProbeError("Excel did not open copy read-only")
            stage = "first export"
            first = audit_vba_components(book)
            stage = "second export without save or changes"
            second = audit_vba_components(book)
            stage = "compare repeated exports"
            try:
                assert_vba_semantics_preserved(
                    first, second, journal_codename=None,
                    expect_new_document=False,
                )
                repeatable = True
                mismatch = "none"
            except VbaSemanticAuditError as exc:
                repeatable = False
                mismatch = str(exc)
            stage = "close without saving"
            book.Close(SaveChanges=False)
            book = None
            app.Quit()
            app = None
        if _file_sha256(source) != source_hash:
            raise RepeatabilityProbeError("Source changed during repeatability audit")
        return {"repeatable": repeatable, "mismatch": mismatch, "components": len(first)}
    except Exception as exc:
        msg = " ".join(str(exc).replace(str(source), "<SOURCE>").split())[:350]
        raise RepeatabilityProbeError(
            f"Stage={stage}; cause={type(exc).__name__}; detail={msg}"
        ) from exc
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


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Isolated, read-only VBA export repeatability check")
    parser.add_argument("--source", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = probe_vba_export_repeatability(args.source)
    except Exception as exc:
        print(f"SAFETY STOP: VBA repeatability diagnosis failed ({type(exc).__name__})")
        print("DIAGNOSTIC: " + " ".join(str(exc).split())[:500])
        return 2
    print("VBA READ-ONLY REPEATABILITY PROBE COMPLETED")
    print(f"Components audited: {result['components']}")
    print("Two exports without saving identical: " + ("YES" if result["repeatable"] else "NO"))
    print("Export comparison: " + result["mismatch"])
    print("Source untouched: YES")
    print("No journal installed; no workbook saved; disposable copy deleted")
    return 0 if result["repeatable"] else 3


if __name__ == "__main__":
    raise SystemExit(main())
