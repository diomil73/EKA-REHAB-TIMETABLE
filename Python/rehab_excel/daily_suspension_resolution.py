from __future__ import annotations

from datetime import date
from pathlib import Path
import sys
import unicodedata
from typing import Iterable

from rehab_core.models import (
    DailySessionCancellation,
    Patient,
    PatientType,
    Session,
    SessionCancellationKind,
)


class DailySuspensionResolutionError(RuntimeError):
    """Raised when unresolved sessions cannot be persisted safely as suspensions."""


SUSPENSION_STATUS = "Αναστολή θεραπείας"
SUSPENSION_REASON = "Δεν βρέθηκε έγκυρη αναπλήρωση από τον global solver"


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value or "").strip().upper())
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def build_unresolved_suspensions(
    unresolved_request_ids: Iterable[str],
    *,
    sessions: Iterable[Session],
    target_date: date,
) -> tuple[DailySessionCancellation, ...]:
    """Convert solver leftovers into explicit date-specific treatment suspensions."""

    session_by_id = {session.session_id: session for session in sessions}
    requested = tuple(dict.fromkeys(str(value).strip() for value in unresolved_request_ids if str(value).strip()))
    suspensions: list[DailySessionCancellation] = []

    for session_id in requested:
        session = session_by_id.get(session_id)
        if session is None:
            raise DailySuspensionResolutionError(
                f"Unresolved session {session_id!r} does not exist in the materialized day"
            )
        if session.session_date != target_date:
            raise DailySuspensionResolutionError(
                f"Unresolved session {session_id!r} belongs to {session.session_date.isoformat()}, "
                f"not {target_date.isoformat()}"
            )
        suspensions.append(
            DailySessionCancellation(
                cancellation_id=f"global-suspension:{session_id}:{target_date.isoformat()}",
                target_session_id=session_id,
                cancellation_date=target_date,
                kind=SessionCancellationKind.TREATMENT_SUSPENDED,
                reason=f"{SUSPENSION_STATUS} | {SUSPENSION_REASON}",
            )
        )

    return tuple(suspensions)


def _patient_display(patient: Patient) -> str:
    if patient.patient_type == PatientType.OUTPATIENT:
        return f"ΕΞ | {patient.display_name}"
    return patient.display_name


def _find_patient_header_row(ws) -> int:
    expected = ("ΑΣΘΕΝΗΣ", "ΟΛΗ ΗΜΕΡΑ", "ΩΡΑ", "ΚΑΤΑΣΤΑΣΗ", "ΣΧΟΛΙΟ")
    for row in range(1, 101):
        current = tuple(_norm(ws.Cells(row, col).Value) for col in range(1, 6))
        if current == expected:
            return row
    raise DailySuspensionResolutionError("Could not locate DAILY_INPUT patient header")


def persist_daily_input_suspensions(
    workbook_path: str | Path,
    suspensions: Iterable[DailySessionCancellation],
    *,
    sessions: Iterable[Session],
    patients: Iterable[Patient],
    max_patient_rows: int = 20,
) -> int:
    """Persist generated suspensions into unused DAILY_INPUT patient rows using Excel COM."""

    suspension_list = tuple(suspensions)
    if not suspension_list:
        return 0
    if max_patient_rows < 1:
        raise ValueError("max_patient_rows must be positive")
    if sys.platform != "win32":
        raise DailySuspensionResolutionError("Windows + Microsoft Excel are required")

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailySuspensionResolutionError("pywin32 is required") from exc

    session_by_id = {session.session_id: session for session in sessions}
    patient_by_id = {patient.patient_id: patient for patient in patients}
    rows: list[tuple[str, str, str]] = []
    for suspension in suspension_list:
        if suspension.kind != SessionCancellationKind.TREATMENT_SUSPENDED:
            raise DailySuspensionResolutionError(
                f"Cancellation {suspension.cancellation_id!r} is not a treatment suspension"
            )
        session = session_by_id.get(suspension.target_session_id)
        if session is None:
            raise DailySuspensionResolutionError(
                f"Suspension target {suspension.target_session_id!r} is missing"
            )
        patient = patient_by_id.get(session.patient_id)
        if patient is None:
            raise DailySuspensionResolutionError(
                f"Patient {session.patient_id!r} is missing for suspension target"
            )
        rows.append(
            (
                _patient_display(patient),
                session.start_time.strftime("%H:%M"),
                SUSPENSION_REASON,
            )
        )

    excel = None
    workbook = None
    try:
        path = Path(workbook_path).resolve()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        ws = workbook.Worksheets("DAILY_INPUT")
        header_row = _find_patient_header_row(ws)
        first_row = header_row + 1
        last_row = first_row + max_patient_rows - 1

        empty_rows = [
            row
            for row in range(first_row, last_row + 1)
            if not str(ws.Cells(row, 1).Value or "").strip()
        ]
        if len(empty_rows) < len(rows):
            raise DailySuspensionResolutionError(
                f"DAILY_INPUT has only {len(empty_rows)} empty patient row(s), "
                f"but {len(rows)} suspension row(s) are required"
            )

        for row_number, (patient_name, time_text, comment) in zip(empty_rows, rows):
            ws.Cells(row_number, 1).Value = patient_name
            ws.Cells(row_number, 2).Value = "ΟΧΙ"
            ws.Cells(row_number, 3).NumberFormat = "@"
            ws.Cells(row_number, 3).Value2 = time_text
            ws.Cells(row_number, 4).Value = SUSPENSION_STATUS
            ws.Cells(row_number, 5).Value = comment

        workbook.Save()
        return len(rows)
    except DailySuspensionResolutionError:
        raise
    except Exception as exc:
        raise DailySuspensionResolutionError(
            f"Could not persist treatment suspensions into DAILY_INPUT: {exc}"
        ) from exc
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
