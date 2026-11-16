"""Restore the established MASTER cell contents after private native DEMO seeding.

MASTER projection construction sorts patients with formula templates. Its
legacy treatment templates display only a *single* PATIENT_PLANNER source
cell, whereas the already-established PatientProgramMasterRender VBA displays
the complete time / weekdays / optional provider on separate lines.

Re-use the existing application macro, never duplicate styling or treatment
layout rules. A private staged DEMO without that installed macro fails closed.
No original workbook, source schedule or historical proof is modified.
"""
from __future__ import annotations

from pathlib import Path
import sys

from openpyxl import load_workbook

from rehab_excel.master_projection import build_master_projection
from rehab_excel.patient_program_master_render_vba import (
    MASTER_RENDER_MODULE_NAME,
    MASTER_RENDER_MODULE_CODE,
)
from rehab_excel.trial_demo_planner_mirror import (
    choose_planner_demo_mirrors,
)
from rehab_excel.trial_demo_source_projection import TrialDemoSourceProjection


# This mapping is the existing MASTER writer/renderer column contract.
# No new columns, merged cells, or formatting rules are introduced.
_MASTER_TREATMENT_COLUMNS = {
    "ΦΘ": 4,
    "Ρομποτικό": 5,
    "Πισίνα": 6,
    "Ανακλινόμενο": 7,
    "Εργο": 8,
    "Λογο": 9,
    "ΕΦΑ": 10,
    "Ψυχ": 11,
}
_PROVIDER_IN_MASTER = frozenset({"ΦΘ", "Ρομποτικό", "Εργο", "Λογο", "Ψυχ"})


class DemoMasterDisplayError(RuntimeError):
    """Existing MASTER rendering could not be retained safely."""


def _expected_lines(mirror) -> str:
    lines = [f"{mirror.start_time:%H:%M}", mirror.day_pattern]
    if mirror.treatment in _PROVIDER_IN_MASTER and mirror.therapist:
        lines.append(mirror.therapist)
    return "\n".join(lines)


def _require_existing_macro(book) -> None:
    try:
        component = book.VBProject.VBComponents(MASTER_RENDER_MODULE_NAME)
        code = str(component.CodeModule.Lines(
            1, int(component.CodeModule.CountOfLines),
        ))
    except Exception:
        raise DemoMasterDisplayError(
            "SAFETY STOP: installed MASTER treatment renderer is missing"
        ) from None
    if (
        "Public Sub RefreshMasterPatientProgram(ByVal patientId As String)" not in code
        or "BuildTreatmentDisplay(planner, plannerRow, treatment)" not in code
        or ".WrapText = True" not in code
    ):
        raise DemoMasterDisplayError(
            "SAFETY STOP: installed MASTER renderer differs from established contract"
        )


def refresh_private_demo_master_display(
    workbook_path: str | Path, projection: TrialDemoSourceProjection,
) -> None:
    """Run only the EXISTING renderer in the isolated staged workbook."""
    if sys.platform != "win32":
        raise DemoMasterDisplayError("Native Windows Excel is required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        raise DemoMasterDisplayError("pywin32 is required") from None
    source = Path(workbook_path).resolve()
    if source.suffix.casefold() != ".xlsm" or "_DEMO_" not in source.name.upper():
        raise DemoMasterDisplayError("Refusing to render a non-DEMO workbook")
    # One instance, one open, one save. No original input or H2-H10 proof.
    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        book = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=False)
        _require_existing_macro(book)
        macro = f"'{book.Name}'!{MASTER_RENDER_MODULE_NAME}.RefreshMasterPatientProgram"
        for record in projection.patient_registry.rows:
            if record[5] != "Εσωτερικός":
                continue
            excel.Run(macro, str(record[0]))
        book.Save()
    except DemoMasterDisplayError:
        raise
    except Exception:
        # Never expose private patient identifiers or local workbook paths.
        raise DemoMasterDisplayError(
            "SAFETY STOP: established MASTER VBA rendering failed"
        ) from None
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def verify_private_demo_master_display(
    workbook_path: str | Path, projection: TrialDemoSourceProjection,
) -> int:
    """Independent disk readback of already-defined time/day/provider cells."""
    expected = choose_planner_demo_mirrors(projection.weekly_slots.rows)
    rows = {
        item.patient_id: item.target_row
        for item in build_master_projection(workbook_path)
    }
    wb = load_workbook(
        workbook_path, read_only=False, data_only=False, keep_vba=True,
    )
    checked = 0
    try:
        ws = wb["MASTER_SCHEDULE"]
        for mirror in expected:
            row = rows.get(mirror.patient_id)
            if row is None:
                # MASTER is inpatient-only. A corresponding outpatient entry
                # belongs to its established outpatient display.
                continue
            col = _MASTER_TREATMENT_COLUMNS[mirror.treatment]
            cell = ws.cell(row, col)
            if (
                cell.value != _expected_lines(mirror)
                or cell.alignment.wrap_text is not True
            ):
                raise DemoMasterDisplayError(
                    "SAFETY STOP: saved MASTER time/weekday/provider display differs"
                )
            checked += 1
    except DemoMasterDisplayError:
        raise
    except Exception:
        raise DemoMasterDisplayError(
            "SAFETY STOP: could not independently verify saved MASTER display"
        ) from None
    finally:
        wb.close()
    return checked
