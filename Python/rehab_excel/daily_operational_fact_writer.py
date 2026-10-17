from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from pathlib import Path
import sys
import unicodedata
from typing import Iterable

from rehab_core.models import DailySessionCancellation, Patient, Session, SessionCancellationKind


DAILY_INPUT_SHEET = "DAILY_INPUT"
TREATMENT_SUSPENSION_STATUS = "Αναστολή θεραπείας"
_PATIENT_HEADER = ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο")


class DailyOperationalFactWriteError(RuntimeError):
    pass


@dataclass(frozen=True)
class TreatmentSuspensionRow:
    target_session_id: str
    patient_name: str
    start_time: time
    status: str = TREATMENT_SUSPENSION_STATUS


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value or "").strip().upper())
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def build_treatment_suspension_rows(
    suspensions: Iterable[DailySessionCancellation],
    *,
    sessions: Iterable[Session],
    patients: Iterable[Patient],
) -> tuple[TreatmentSuspensionRow, ...]:
    session_by_id = {session.session_id: session for session in sessions}
    patient_by_id = {patient.patient_id: patient for patient in patients}
    rows: list[TreatmentSuspensionRow] = []

    for suspension in suspensions:
        if suspension.kind != SessionCancellationKind.TREATMENT_SUSPENDED:
            raise DailyOperationalFactWriteError(
                f"Cancellation {suspension.cancellation_id!r} is not treatment suspension"
            )
        try:
            session = session_by_id[suspension.target_session_id]
        except KeyError as exc:
            raise DailyOperationalFactWriteError(
                f"Unknown target session {suspension.target_session_id!r}"
            ) from exc
        try:
            patient = patient_by_id[session.patient_id]
        except KeyError as exc:
            raise DailyOperationalFactWriteError(
                f"Unknown patient {session.patient_id!r} for suspension"
            ) from exc
        if suspension.cancellation_date != session.session_date:
            raise DailyOperationalFactWriteError(
                f"Suspension date {suspension.cancellation_date} does not match "
                f"session date {session.session_date}"
            )
        rows.append(
            TreatmentSuspensionRow(
                target_session_id=session.session_id,
                patient_name=patient.display_name,
                start_time=session.start_time,
            )
        )

    return tuple(rows)


def _find_patient_header_row(ws) -> int:
    wanted = tuple(_norm(value) for value in _PATIENT_HEADER)
    for row in range(1, 101):
        current = tuple(_norm(ws.Cells(row, col).Value) for col in range(1, 6))
        if current == wanted:
            return row
    raise DailyOperationalFactWriteError(
        "Could not locate DAILY_INPUT patient header for operational fact persistence"
    )


def _time_text(value: object) -> str:
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%H:%M")
        except Exception:
            pass
    text = str(value or "").strip()
    if len(text) >= 5 and text[2] == ":":
        return text[:5]
    return text


def _existing_suspension_keys(ws, header_row: int) -> set[tuple[str, str, str]]:
    keys: set[tuple[str, str, str]] = set()
    last_row = max(int(ws.UsedRange.Rows.Count), header_row + 40)
    for row in range(header_row + 1, last_row + 1):
        patient = str(ws.Cells(row, 1).Value or "").strip()
        status = str(ws.Cells(row, 4).Value or "").strip()
        if not patient or _norm(status) != _norm(TREATMENT_SUSPENSION_STATUS):
            continue
        keys.add(
            (
                _norm(patient.removeprefix("ΕΞ |").strip()),
                _time_text(ws.Cells(row, 3).Value),
                _norm(status),
            )
        )
    return keys


def _first_blank_patient_row(ws, header_row: int) -> int:
    row = header_row + 1
    ceiling = header_row + 500
    while row <= ceiling:
        if not str(ws.Cells(row, 1).Value or "").strip():
            return row
        row += 1
    raise DailyOperationalFactWriteError("DAILY_INPUT has no blank patient row for suspension")


def append_treatment_suspensions_to_daily_input(
    workbook_path: str | Path,
    rows: Iterable[TreatmentSuspensionRow],
) -> int:
    """Persist suspension facts into the output workbook's DAILY_INPUT sheet.

    The function is idempotent for patient/time/status. It intentionally writes
    only to the provided workbook path, which should be a preview/output copy.
    """

    row_list = tuple(rows)
    if not row_list:
        return 0
    if sys.platform != "win32":
        raise DailyOperationalFactWriteError(
            "Treatment suspension persistence requires Windows + Microsoft Excel"
        )
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DailyOperationalFactWriteError("pywin32 is required") from exc

    path = Path(workbook_path).resolve()
    excel = None
    workbook = None
    written = 0
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        try:
            ws = workbook.Worksheets(DAILY_INPUT_SHEET)
        except Exception as exc:
            raise DailyOperationalFactWriteError(
                f"Workbook has no {DAILY_INPUT_SHEET} sheet"
            ) from exc

        header_row = _find_patient_header_row(ws)
        existing = _existing_suspension_keys(ws, header_row)

        for fact in row_list:
            key = (_norm(fact.patient_name), fact.start_time.strftime("%H:%M"), _norm(fact.status))
            if key in existing:
                continue
            target_row = _first_blank_patient_row(ws, header_row)
            ws.Cells(target_row, 1).Value = fact.patient_name
            ws.Cells(target_row, 2).Value = "ΟΧΙ"
            try:
                ws.Cells(target_row, 3).NumberFormat = "@"
            except Exception:
                pass
            ws.Cells(target_row, 3).Value2 = fact.start_time.strftime("%H:%M")
            ws.Cells(target_row, 4).Value = fact.status
            ws.Cells(target_row, 5).Value = "Ορφανός - ημερήσια επιχειρησιακή απόφαση"
            existing.add(key)
            written += 1

        workbook.Save()
    except DailyOperationalFactWriteError:
        raise
    except Exception as exc:
        raise DailyOperationalFactWriteError(
            f"Could not persist treatment suspension facts: {exc}"
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

    return written
