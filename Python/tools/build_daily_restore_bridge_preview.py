from __future__ import annotations

import argparse
from datetime import date, datetime
from hashlib import sha256
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_excel.daily_change_ledger_vba import (  # noqa: E402
    LEDGER_SHEET,
    MODULE_NAME,
    VBA_CODE,
)
from rehab_excel.daily_restore_in_session_vba import build_in_session_ledger_vba  # noqa: E402
from rehab_excel.daily_replacement_facts import (  # noqa: E402
    ACTIVE_STATUS,
    read_daily_replacement_fact_rows,
)
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402


class DailyRestoreBridgePreviewError(RuntimeError):
    pass


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _has_vba(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            return "xl/vbaProject.bin" in archive.namelist()
    except Exception:
        return False


def _remove_component_if_present(vbproject, name: str) -> None:
    try:
        component = vbproject.VBComponents(name)
    except Exception:
        return
    vbproject.VBComponents.Remove(component)


def _read_daily_date(path: Path) -> date:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        if "DAILY_INPUT" not in wb.sheetnames:
            raise DailyRestoreBridgePreviewError("Workbook has no DAILY_INPUT sheet")
        raw = wb["DAILY_INPUT"]["B2"].value
    finally:
        wb.close()

    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    text = str(raw or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise DailyRestoreBridgePreviewError(f"Invalid DAILY_INPUT date: {raw!r}")


def _authoritative_ledger_rows(path: Path):
    """Build display metadata from durable replacement facts, never from ledger state."""

    target_date = _read_daily_date(path)
    facts = read_daily_replacement_fact_rows(
        path,
        target_date=target_date,
        include_restored=True,
    )
    if not facts:
        raise DailyRestoreBridgePreviewError(
            "Workbook has no authoritative daily replacement facts for the DAILY_INPUT date"
        )

    patients = read_patient_registry(path)
    base_entries = read_unified_base_schedule(path)
    sessions = materialize_sessions_for_date(base_entries, target_date)
    patient_by_id = {patient.patient_id: patient for patient in patients}
    session_by_id = {session.session_id: session for session in sessions}

    rows: list[tuple[object, ...]] = []
    active_count = 0
    for fact in facts:
        session = session_by_id.get(fact.target_session_id)
        if session is None:
            raise DailyRestoreBridgePreviewError(
                f"Replacement fact {fact.replacement_id!r} references unknown session "
                f"{fact.target_session_id!r}"
            )
        patient = patient_by_id.get(fact.patient_id)
        patient_name = patient.display_name if patient is not None else fact.patient_id
        if fact.status == ACTIVE_STATUS:
            active_count += 1
        rows.append(
            (
                fact.replacement_id,
                datetime.combine(fact.replacement_date, datetime.min.time()),
                fact.patient_id,
                patient_name,
                fact.original_therapist_id,
                session.start_time.strftime("%H:%M"),
                fact.replacement_therapist_id,
                fact.replacement_time.strftime("%H:%M"),
                fact.reason or "temporary replacement",
                fact.status,
            )
        )

    if active_count == 0:
        raise DailyRestoreBridgePreviewError(
            "Authoritative replacement facts have no ACTIVE rows for bridge smoke testing"
        )
    return tuple(rows), active_count


def _sync_ledger_from_authoritative_facts(workbook, rows) -> None:
    """Make the UI ledger a projection of the authoritative fact sheet.

    Existing snapshot/audit columns are retained by immutable ReplacementId when
    available. Missing snapshot fields are intentionally left blank rather than
    inventing history. Columns 1-9 and 12 are always rebuilt from durable facts.
    """

    try:
        ledger = workbook.Worksheets(LEDGER_SHEET)
    except Exception as exc:
        raise DailyRestoreBridgePreviewError(
            f"Source copy has no {LEDGER_SHEET} sheet"
        ) from exc

    last_row = int(ledger.Cells(ledger.Rows.Count, 1).End(-4162).Row)  # xlUp
    preserved: dict[str, tuple[object, ...]] = {}
    if last_row >= 2:
        for row_index in range(2, last_row + 1):
            replacement_id = str(ledger.Cells(row_index, 1).Value or "").strip()
            if not replacement_id:
                continue
            preserved[replacement_id] = tuple(
                ledger.Cells(row_index, col).Value for col in range(10, 19)
            )
        ledger.Range(ledger.Cells(2, 1), ledger.Cells(last_row, 19)).ClearContents()

    for row_index, row in enumerate(rows, start=2):
        replacement_id = str(row[0])
        for col, value in enumerate(row[:9], start=1):
            ledger.Cells(row_index, col).Value = value
        ledger.Cells(row_index, 12).Value = row[9]

        old = preserved.get(replacement_id)
        if old is not None:
            for offset, value in enumerate(old, start=10):
                if offset == 12:
                    continue
                ledger.Cells(row_index, offset).Value = value


def build_daily_restore_bridge_preview(
    source_path: str | Path,
    output_path: str | Path,
    *,
    overwrite: bool = False,
) -> tuple[Path, int]:
    """Clone a trusted daily workbook and install a fact-synchronized ledger bridge.

    DAILY_REPLACEMENT_FACTS is authoritative. DAILY_CHANGE_LEDGER is rebuilt as
    its date-scoped UI projection before the current VBA bridge module is installed.
    This is essential for rerun workbooks where older ACTIVE/RESTORED facts predate
    the most recent ledger-render cycle.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()

    if not source.exists():
        raise DailyRestoreBridgePreviewError(f"Source workbook not found: {source}")
    if source == output:
        raise DailyRestoreBridgePreviewError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise DailyRestoreBridgePreviewError("Source and output must both be .xlsm files")
    if not _has_vba(source):
        raise DailyRestoreBridgePreviewError("Source workbook has no VBA project")
    if output.exists() and not overwrite:
        raise DailyRestoreBridgePreviewError(f"Output already exists: {output}")
    if sys.platform != "win32":
        raise DailyRestoreBridgePreviewError("Windows with Microsoft Excel is required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyRestoreBridgePreviewError("pywin32 is required") from exc

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        try:
            output.unlink()
        except PermissionError as exc:
            raise DailyRestoreBridgePreviewError(
                f"Output workbook is open or locked by Excel: {output}"
            ) from exc

    shutil.copy2(source, output)
    excel = None
    workbook = None
    failure: Exception | None = None
    fact_rows: tuple[tuple[object, ...], ...] = ()
    active_count = 0
    try:
        fact_rows, active_count = _authoritative_ledger_rows(output)
        effective_vba = build_in_session_ledger_vba(VBA_CODE)

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False)
        _sync_ledger_from_authoritative_facts(workbook, fact_rows)

        vbproject = workbook.VBProject
        _remove_component_if_present(vbproject, MODULE_NAME)
        module = vbproject.VBComponents.Add(1)
        module.Name = MODULE_NAME
        module.CodeModule.AddFromString(effective_vba)

        workbook.Save()
        excel.Run(f"'{workbook.Name}'!RefreshDailyChangeLedgerView")
        workbook.Save()
    except Exception as exc:
        failure = exc
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
        output.unlink(missing_ok=True)
        if isinstance(failure, DailyRestoreBridgePreviewError):
            raise failure
        raise DailyRestoreBridgePreviewError(str(failure)) from failure

    if _sha256(source) != source_before:
        output.unlink(missing_ok=True)
        raise DailyRestoreBridgePreviewError("Source workbook changed during preview build")
    if not _has_vba(output):
        output.unlink(missing_ok=True)
        raise DailyRestoreBridgePreviewError("Preview lost its VBA project")

    return output, active_count


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Create a separate native Excel smoke-test copy of an authoritative daily "
            "workbook, synchronize DAILY_CHANGE_LEDGER from durable replacement facts, "
            "and refresh modDailyChangeLedger to the current in-session restore bridge."
        )
    )
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--overwrite", action="store_true")
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    try:
        output, active_count = build_daily_restore_bridge_preview(
            args.source,
            args.output,
            overwrite=args.overwrite,
        )
    except DailyRestoreBridgePreviewError as exc:
        print(f"SAFETY STOP: {exc}")
        return 2

    print("DAILY RESTORE BRIDGE PREVIEW OK")
    print(f"Source: {args.source.resolve()}")
    print(f"Preview: {output}")
    print(f"Active ledger rows: {active_count}")
    print("Ledger synchronized from authoritative facts: True")
    print("Source unchanged: True")
    print("VBA preserved: True")
    print("Ledger bridge module refreshed: True")
    print("Restore execution mode: hidden Python + in-session Excel apply")
    print("NEXT: open only this preview and exercise the intended Excel restore control.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())