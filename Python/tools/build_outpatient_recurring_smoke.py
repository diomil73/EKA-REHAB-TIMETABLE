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

from rehab_excel.outpatient_conflict_guard_vba import (  # noqa: E402
    install_outpatient_conflict_guard,
)
from rehab_excel.outpatient_schedule_edit_vba import (  # noqa: E402
    EDIT_MACRO_NAME,
    install_outpatient_schedule_edit_support,
)
from rehab_excel.outpatient_schedule_form_vba import (  # noqa: E402
    FORM_CODE,
    FORM_NAME,
    MACRO_NAME,
    MODULE_CODE,
    MODULE_NAME,
)
from rehab_excel.outpatient_schedule_manager_vba import (  # noqa: E402
    FORM_NAME as MANAGER_FORM_NAME,
    MACRO_NAME as MANAGER_MACRO_NAME,
    MODULE_NAME as MANAGER_MODULE_NAME,
    install_outpatient_schedule_manager,
)


class OutpatientRecurringSmokeBuildError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _safe_unlink(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _component_if_present(vbproject, name: str):
    try:
        return vbproject.VBComponents(name)
    except Exception:
        return None


def _replace_component_code(component, code: str) -> None:
    module = component.CodeModule
    count = int(getattr(module, "CountOfLines", 0) or 0)
    if count > 0:
        module.DeleteLines(1, count)
    module.AddFromString(code)


def _ensure_form_control(designer, prog_id: str, name: str):
    try:
        return designer.Controls(name)
    except Exception:
        return designer.Controls.Add(prog_id, name, True)


def _install_schedule_form(vbproject) -> None:
    """Install or refresh the outpatient form without remove/re-add COM churn.

    Reusing an existing VBComponent is important in the combined one-session smoke
    build. Excel/VBIDE can reject immediate remove-and-recreate operations for a
    UserForm with the generic 0x800A03EC COM error even though the component was
    just removed successfully.
    """

    module = _component_if_present(vbproject, MODULE_NAME)
    if module is None:
        module = vbproject.VBComponents.Add(1)
        module.Name = MODULE_NAME
    _replace_component_code(module, MODULE_CODE)

    form = _component_if_present(vbproject, FORM_NAME)
    if form is None:
        form = vbproject.VBComponents.Add(3)
        form.Name = FORM_NAME

    designer = form.Designer
    designer.Caption = "Πρόγραμμα εξωτερικού ασθενή"
    for prog_id, name, caption, top in (
        ("Forms.Label.1", "lblTitle", "", 15),
        ("Forms.Label.1", "lblPatient", "", 48),
        ("Forms.ComboBox.1", "cboPatient", "", 44),
        ("Forms.Label.1", "lblTreatment", "", 92),
        ("Forms.ComboBox.1", "cboTreatment", "", 88),
        ("Forms.Label.1", "lblTime", "", 136),
        ("Forms.ComboBox.1", "cboTime", "", 132),
        ("Forms.Label.1", "lblDays", "", 180),
        ("Forms.ComboBox.1", "cboDays", "", 176),
        ("Forms.Label.1", "lblTherapist", "", 224),
        ("Forms.ComboBox.1", "cboTherapist", "", 220),
        ("Forms.Label.1", "lblSuggestMode", "", 268),
        ("Forms.ComboBox.1", "cboSuggestMode", "", 288),
        ("Forms.CommandButton.1", "cmdSuggest", "Βρες εναλλακτικές", 323),
        ("Forms.ListBox.1", "lstSuggestions", "", 365),
        ("Forms.CommandButton.1", "cmdUseSuggestion", "Χρήση επιλογής", 470),
        ("Forms.Label.1", "lblInfo", "", 505),
        ("Forms.CommandButton.1", "cmdCancel", "Ακύρωση", 550),
        ("Forms.CommandButton.1", "cmdSave", "Έλεγχος και preview", 550),
    ):
        control = _ensure_form_control(designer, prog_id, name)
        if caption:
            control.Caption = caption
        control.Left = 24
        control.Top = top

    _replace_component_code(form, FORM_CODE)


def _verify_components(vbproject) -> None:
    names = {
        vbproject.VBComponents(index).Name
        for index in range(1, vbproject.VBComponents.Count + 1)
    }
    required = {
        FORM_NAME,
        MODULE_NAME,
        "modOutpatientScheduleEdit",
        MANAGER_FORM_NAME,
        MANAGER_MODULE_NAME,
        "frmSchedulingConflict",
        "modSchedulingConflict",
        "modOutpatientInSession",
    }
    missing = sorted(required - names)
    if missing:
        raise OutpatientRecurringSmokeBuildError(
            "Outpatient recurring smoke components missing: " + ", ".join(missing)
        )


def build_outpatient_recurring_smoke(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> Path:
    """Build create/edit/manager/suggestion/conflict support in one Excel session.

    The trusted source is copied byte-for-byte first and is never opened for
    mutation. All VBA installation then happens while the final smoke copy is
    open once, eliminating cross-process stage-file lock races.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise OutpatientRecurringSmokeBuildError(f"Source workbook not found: {source}")
    if source == output:
        raise OutpatientRecurringSmokeBuildError("Output must differ from source")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise OutpatientRecurringSmokeBuildError("Source and output must both be .xlsm")
    if output.exists() and not overwrite:
        raise OutpatientRecurringSmokeBuildError(f"Output already exists: {output}")
    if sys.platform != "win32":
        raise OutpatientRecurringSmokeBuildError("Windows Excel is required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise OutpatientRecurringSmokeBuildError("pywin32 is required") from exc

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except OSError as exc:
            raise OutpatientRecurringSmokeBuildError(
                f"Existing smoke workbook is open or locked: {output}"
            ) from exc
    shutil.copy2(source, output)

    excel = None
    workbook = None
    failure: tuple[str, Exception] | None = None
    phase = "Excel startup"
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        phase = "workbook open"
        workbook = excel.Workbooks.Open(
            str(output), UpdateLinks=0, ReadOnly=False
        )
        vbproject = workbook.VBProject

        phase = "schedule form install"
        _install_schedule_form(vbproject)

        phase = "edit support install"
        install_outpatient_schedule_edit_support(vbproject)

        phase = "schedule manager install"
        install_outpatient_schedule_manager(vbproject)

        phase = "conflict guard install"
        install_outpatient_conflict_guard(vbproject)

        phase = "component verification"
        _verify_components(vbproject)

        phase = "workbook save"
        workbook.Save()
    except Exception as exc:
        failure = (phase, exc)
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

    if failure is not None:
        _safe_unlink(output)
        failed_phase, exc = failure
        if isinstance(exc, OutpatientRecurringSmokeBuildError):
            raise OutpatientRecurringSmokeBuildError(
                f"{failed_phase}: {exc}"
            ) from exc
        raise OutpatientRecurringSmokeBuildError(
            f"Excel outpatient recurring smoke installation failed during {failed_phase}: {exc}"
        ) from exc

    if _sha256(source) != source_before:
        _safe_unlink(output)
        raise OutpatientRecurringSmokeBuildError(
            "Source workbook changed during outpatient recurring smoke build"
        )
    if not output.exists():
        raise OutpatientRecurringSmokeBuildError("Smoke workbook was not produced")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Build one fresh XLSM smoke copy with outpatient recurring create/edit, "
            "management UI, suggestions and conflict popup support."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    try:
        output = build_outpatient_recurring_smoke(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except OutpatientRecurringSmokeBuildError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("OUTPATIENT RECURRING SMOKE BUILD OK")
    print(f"Source: {args.source.resolve()}")
    print(f"Smoke: {output}")
    print(f"Create macro: {MACRO_NAME}")
    print(f"Manage/edit macro: {MANAGER_MACRO_NAME}")
    print(f"Legacy edit alias: {EDIT_MACRO_NAME}")
    print("Source unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
