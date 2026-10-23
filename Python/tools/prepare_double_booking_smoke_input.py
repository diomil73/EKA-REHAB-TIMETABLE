from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_excel.application_contract import assert_application_contract  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_patients import audit_daily_input_patients  # noqa: E402
from rehab_excel.daily_input_sheet import (  # noqa: E402
    build_daily_input_spec,
    create_daily_input_preview,
)
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402


class DoubleBookingSmokeError(RuntimeError):
    pass


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value.strip())
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Date must be YYYY-MM-DD") from exc


def choose_double_booking_smoke_case(sessions, *, therapist_names: tuple[str, ...]):
    """Pick one real target and one real therapist already occupied at the same time."""
    allowed = set(therapist_names)
    by_time: dict[object, list] = {}
    for session in sessions:
        if session.therapist_id in allowed:
            by_time.setdefault(session.start_time, []).append(session)

    candidates = []
    for slot, slot_sessions in by_time.items():
        providers = sorted({item.therapist_id for item in slot_sessions})
        if len(providers) < 2:
            continue
        for target in sorted(slot_sessions, key=lambda x: (x.therapist_id, x.patient_id, x.session_id)):
            busy = next(
                (item for item in slot_sessions if item.therapist_id != target.therapist_id),
                None,
            )
            if busy is not None:
                candidates.append((target, busy))

    if not candidates:
        raise DoubleBookingSmokeError(
            "Could not find a date/time with two different therapists already scheduled"
        )
    return candidates[0]


def _find_header_row(ws, label: str) -> int:
    wanted = label.casefold()
    for row in range(1, 101):
        if str(ws.Cells(row, 1).Value or "").strip().casefold() == wanted:
            return row
    raise DoubleBookingSmokeError(f"Could not locate DAILY_INPUT header: {label}")


def _ensure_therapist_row_capacity(
    ws,
    *,
    therapist_header: int,
    patient_header: int,
    required_rows: int,
) -> int:
    """Expand only the disposable smoke copy when the template has too few rows."""

    available_rows = patient_header - therapist_header - 2
    shortage = max(0, required_rows - available_rows)
    if shortage == 0:
        return patient_header

    insert_start = patient_header - 1
    insert_end = insert_start + shortage - 1
    ws.Rows(f"{insert_start}:{insert_end}").Insert()

    template_row = therapist_header + 1
    try:
        ws.Rows(template_row).Copy()
        ws.Rows(f"{insert_start}:{insert_end}").PasteSpecial(Paste=-4122)
    except Exception:
        pass
    try:
        ws.Application.CutCopyMode = False
    except Exception:
        pass

    return patient_header + shortage


def _write_smoke_rows(
    workbook_path: Path,
    *,
    original_therapist: str,
    busy_therapist: str,
    active_day_therapists: tuple[str, ...],
    patient_name: str,
    original_time,
    timeslots,
) -> None:
    if sys.platform != "win32":
        raise DoubleBookingSmokeError("Windows + Microsoft Excel are required")
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DoubleBookingSmokeError("pywin32 is required") from exc

    excel = workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False)
        ws = workbook.Worksheets("DAILY_INPUT")
        therapist_header = _find_header_row(ws, "Θεραπευτής")
        patient_header = _find_header_row(ws, "Ασθενής")

        # Keep only the deliberately busy therapist active. Every other therapist
        # must be a provider that actually has at least one materialized session
        # on the target date, because DAILY_INPUT validates names against that set.
        absent_ids = tuple(
            therapist_id
            for therapist_id in active_day_therapists
            if therapist_id != busy_therapist
        )
        if original_therapist not in absent_ids:
            raise DoubleBookingSmokeError(
                "Original therapist is not present in target-date materialized sessions"
            )
        patient_header = _ensure_therapist_row_capacity(
            ws,
            therapist_header=therapist_header,
            patient_header=patient_header,
            required_rows=len(absent_ids),
        )
        for offset, therapist_id in enumerate(absent_ids, start=1):
            row = therapist_header + offset
            ws.Cells(row, 1).Value = therapist_id
            ws.Cells(row, 2).Value = "ΝΑΙ"
            ws.Range(ws.Cells(row, 3), ws.Cells(row, 4)).ClearContents()
            ws.Cells(row, 5).Value = "Double-booking smoke"
            ws.Cells(row, 6).Value = "Normal replacement intentionally blocked"

        # Keep the target patient available only at the original time.
        row = patient_header + 1
        for slot in timeslots:
            if slot == original_time:
                continue
            ws.Cells(row, 1).Value = patient_name
            ws.Cells(row, 2).Value = "ΟΧΙ"
            ws.Cells(row, 3).Value = slot.strftime("%H:%M")
            ws.Cells(row, 4).Value = "ΜΗ ΔΙΑΘΕΣΙΜΟΣ"
            ws.Cells(row, 5).Value = "Double-booking smoke"
            row += 1

        workbook.Save()
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a full-app unresolved double-booking smoke input.")
    parser.add_argument("--date", type=_parse_date, required=True)
    parser.add_argument(
        "--source",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "OUTPATIENT_RECURRING_SMOKE.xlsm",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=REPO_ROOT / "Excel" / "previews" / "DOUBLE_BOOKING_SMOKE_INPUT.xlsm",
    )
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    if not source.exists():
        print(f"SAFETY STOP: source workbook not found: {source}")
        return 2

    source_hash = file_sha256(source)
    try:
        source_contract = assert_application_contract(source)
        patient_audit = audit_daily_input_patients(source)
        if not patient_audit.ok:
            raise DoubleBookingSmokeError("Patient identity audit failed")
        settings = read_settings(source)
        base_entries = read_unified_base_schedule(source)
        sessions = materialize_sessions_for_date(base_entries, args.date)
        active_day_therapists = tuple(sorted({session.therapist_id for session in sessions}))
        target, busy = choose_double_booking_smoke_case(
            sessions,
            therapist_names=tuple(settings.therapist_names),
        )
        patients = read_patient_registry(source)
        patient_by_id = {patient.patient_id: patient.display_name for patient in patients}
        patient_name = patient_by_id.get(target.patient_id)
        if not patient_name:
            raise DoubleBookingSmokeError("Target patient identity not found")

        spec = build_daily_input_spec(
            target_date=args.date,
            therapist_names=settings.therapist_names,
            patient_names=patient_audit.dropdown_names,
            patient_statuses=settings.patient_statuses,
            timeslots=settings.standard_timeslots,
        )
        preview, _ = create_daily_input_preview(source, output, spec, overwrite=args.overwrite)
        _write_smoke_rows(
            preview,
            original_therapist=target.therapist_id,
            busy_therapist=busy.therapist_id,
            active_day_therapists=active_day_therapists,
            patient_name=patient_name,
            original_time=target.start_time,
            timeslots=tuple(settings.standard_timeslots),
        )
        output_contract = assert_application_contract(output)
    except (ValueError, KeyError, DoubleBookingSmokeError) as exc:
        try:
            output.unlink(missing_ok=True)
        except OSError:
            pass
        print(f"SAFETY STOP: {exc}")
        return 2

    source_unchanged = file_sha256(source) == source_hash
    print("DOUBLE BOOKING SMOKE INPUT OK")
    print(f"Date: {args.date.isoformat()}")
    print(f"Target patient: {patient_name}")
    print(f"Original therapist: {target.therapist_id}")
    print(f"Original time: {target.start_time.strftime('%H:%M')}")
    print(f"Busy exception therapist: {busy.therapist_id}")
    print(f"Busy therapist existing patient: {patient_by_id.get(busy.patient_id, busy.patient_id)}")
    print(f"Preview: {output}")
    print(f"Source unchanged: {source_unchanged}")
    print(f"Application contract source: OK ({len(source_contract.component_names)} components)")
    print(f"Application contract output: OK ({len(output_contract.component_names)} components)")
    return 0 if source_unchanged else 3


if __name__ == "__main__":
    raise SystemExit(main())
