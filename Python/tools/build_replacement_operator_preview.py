from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import sys

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.replacement_operator_vba import (  # noqa: E402
    FORM_NAME,
    MODULE_NAME,
    install_replacement_operator,
)


class ReplacementOperatorPreviewError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _component_present(vbproject, name: str) -> bool:
    try:
        _ = vbproject.VBComponents(name)
        return True
    except Exception:
        return False


def build_preview(source: Path, output: Path, *, overwrite: bool) -> None:
    if sys.platform != "win32":
        raise ReplacementOperatorPreviewError("This preview builder requires Windows and Excel")
    if not source.exists():
        raise ReplacementOperatorPreviewError(f"Source workbook not found: {source}")
    if source.resolve() == output.resolve():
        raise ReplacementOperatorPreviewError("Output must differ from source")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise ReplacementOperatorPreviewError("Source and output must both be .xlsm")
    if output.exists() and not overwrite:
        raise ReplacementOperatorPreviewError(f"Output already exists: {output}")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ReplacementOperatorPreviewError("pywin32 is required") from exc

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    shutil.copy2(source, output)

    excel = None
    workbook = None
    stage = "starting Excel"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        stage = "opening preview copy"
        workbook = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=False)
        stage = "accessing VBA project"
        vbproject = workbook.VBProject
        _ = vbproject.VBComponents.Count

        stage = "installing replacement operator overlay"
        install_replacement_operator(vbproject)

        stage = "verifying installed route"
        if not _component_present(vbproject, MODULE_NAME):
            raise ReplacementOperatorPreviewError("Replacement operator module was not installed")
        if not _component_present(vbproject, FORM_NAME):
            raise ReplacementOperatorPreviewError("Replacement operator form was not installed")
        daily_module = vbproject.VBComponents("modDailyInputAction").CodeModule
        daily_source = str(daily_module.Lines(1, daily_module.CountOfLines))
        if "If RunReplacementOperatorWorkflow() Then" not in daily_source:
            raise ReplacementOperatorPreviewError("DAILY_INPUT was not routed to operator workflow")
        if "If ResolveDailyReplacementsInPlace() Then" in daily_source:
            raise ReplacementOperatorPreviewError("Legacy automatic continuation is still active")

        stage = "saving preview"
        workbook.Save()
    except ReplacementOperatorPreviewError:
        raise
    except Exception as exc:
        raise ReplacementOperatorPreviewError(f"V16 build failed during {stage}: {exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

    if _sha256(source) != source_before:
        try:
            output.unlink()
        except OSError:
            pass
        raise ReplacementOperatorPreviewError("Source workbook changed during V16 build")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build an isolated operator-controlled replacement preview from V15."
    )
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    try:
        build_preview(source, output, overwrite=args.overwrite)
    except (ReplacementOperatorPreviewError, OSError) as exc:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        print(f"SAFETY STOP: {exc}")
        return 2

    print("REPLACEMENT OPERATOR V16 PREVIEW OK")
    print(f"Preview: {output}")
    print("Source unchanged: True")
    print("Operator module/form present: True")
    print("DAILY_INPUT continuation routed to operator workflow: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
