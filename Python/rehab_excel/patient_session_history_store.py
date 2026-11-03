from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook

from rehab_core.daily_state import DailySessionStatus
from rehab_core.patient_session_history import PatientSessionHistoryRow
from rehab_excel.finalized_day_snapshots import read_finalized_day_snapshots


SHEET_NAME = "PATIENT_SESSION_HISTORY"
HEADERS = (
    "SnapshotId",
    "ServiceDate",
    "SnapshotVersion",
    "SessionId",
    "PatientId",
    "Status",
    "OriginalTherapistId",
    "OriginalTime",
    "EffectiveTherapistId",
    "EffectiveTime",
    "ReplacementId",
    "CancellationId",
    "CancellationKind",
    "Reason",
)


class PatientSessionHistoryStoreError(RuntimeError):
    """Raised when durable patient/session history is malformed or unsafe."""


def _parse_date(value: object, *, row: int) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value or "").strip())
    except ValueError as exc:
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: invalid ServiceDate {value!r}"
        ) from exc


def _parse_time(value: object, *, row: int, field: str, optional: bool = False) -> time | None:
    if value is None or str(value).strip() == "":
        if optional:
            return None
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: {field} is blank"
        )
    if isinstance(value, datetime):
        return value.time().replace(tzinfo=None)
    if isinstance(value, time):
        return value.replace(tzinfo=None)
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            pass
    raise PatientSessionHistoryStoreError(
        f"{SHEET_NAME} row {row}: invalid {field} {value!r}"
    )


def _required_int(value: object, *, row: int, field: str) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: invalid {field} {value!r}"
        ) from exc
    if result < 1:
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: {field} must be at least 1"
        )
    return result


def _required_text(value: object, *, row: int, field: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: {field} is blank"
        )
    return text


def _optional_text(value: object) -> str | None:
    text = str(value or "").strip()
    return text or None


def _row_from_values(values: tuple[object, ...], *, row: int) -> PatientSessionHistoryRow:
    status_text = _required_text(values[5], row=row, field="Status").casefold()
    try:
        status = DailySessionStatus(status_text)
    except ValueError as exc:
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: unsupported Status {values[5]!r}"
        ) from exc
    if status == DailySessionStatus.THERAPIST_ABSENT:
        raise PatientSessionHistoryStoreError(
            f"{SHEET_NAME} row {row}: finalized history cannot contain unresolved therapist absence"
        )

    return PatientSessionHistoryRow(
        snapshot_id=_required_text(values[0], row=row, field="SnapshotId"),
        service_date=_parse_date(values[1], row=row),
        snapshot_version=_required_int(values[2], row=row, field="SnapshotVersion"),
        session_id=_required_text(values[3], row=row, field="SessionId"),
        patient_id=_required_text(values[4], row=row, field="PatientId"),
        status=status,
        original_therapist_id=_required_text(values[6], row=row, field="OriginalTherapistId"),
        original_time=_parse_time(values[7], row=row, field="OriginalTime"),
        effective_therapist_id=_optional_text(values[8]),
        effective_time=_parse_time(values[9], row=row, field="EffectiveTime", optional=True),
        replacement_id=_optional_text(values[10]),
        cancellation_id=_optional_text(values[11]),
        cancellation_kind=_optional_text(values[12]),
        reason=_optional_text(values[13]),
    )


def read_patient_session_history_rows(
    workbook_path: str | Path,
    *,
    snapshot_id: str | None = None,
    service_date: date | None = None,
    patient_id: str | None = None,
) -> tuple[PatientSessionHistoryRow, ...]:
    """Read immutable patient/session rows without mutating the workbook."""

    path = Path(workbook_path)
    wb = load_workbook(
        path,
        read_only=True,
        data_only=False,
        keep_vba=path.suffix.casefold() == ".xlsm",
    )
    try:
        if SHEET_NAME not in wb.sheetnames:
            return ()
        ws = wb[SHEET_NAME]
        actual_headers = tuple(
            str(ws.cell(1, col).value or "").strip()
            for col in range(1, len(HEADERS) + 1)
        )
        if actual_headers != HEADERS:
            raise PatientSessionHistoryStoreError(
                f"{SHEET_NAME} headers do not match authoritative schema: {actual_headers!r}"
            )

        result: list[PatientSessionHistoryRow] = []
        seen: set[tuple[str, str]] = set()
        patient_filter = patient_id.casefold() if patient_id is not None else None
        for row_number, row_values in enumerate(
            ws.iter_rows(min_row=2, min_col=1, max_col=len(HEADERS), values_only=True),
            start=2,
        ):
            values = tuple(row_values)
            if not any(value is not None and str(value).strip() for value in values):
                continue
            parsed = _row_from_values(values, row=row_number)
            key = (parsed.snapshot_id.casefold(), parsed.session_id.casefold())
            if key in seen:
                raise PatientSessionHistoryStoreError(
                    f"Duplicate patient/session history key {key!r}"
                )
            seen.add(key)
            if snapshot_id is not None and parsed.snapshot_id != snapshot_id:
                continue
            if service_date is not None and parsed.service_date != service_date:
                continue
            if patient_filter is not None and parsed.patient_id.casefold() != patient_filter:
                continue
            result.append(parsed)
        return tuple(result)
    finally:
        wb.close()


def read_current_patient_session_history_rows(
    workbook_path: str | Path,
    *,
    service_date: date | None = None,
    patient_id: str | None = None,
) -> tuple[PatientSessionHistoryRow, ...]:
    """Read only rows belonging to CURRENT_FINAL day snapshots.

    Superseded session rows remain durable for audit but are excluded from
    operational patient statistics.
    """

    current = read_finalized_day_snapshots(
        workbook_path,
        service_date=service_date,
        current_only=True,
    )
    current_ids = {snapshot.snapshot_id for snapshot in current}
    if not current_ids:
        return ()
    rows = read_patient_session_history_rows(
        workbook_path,
        service_date=service_date,
        patient_id=patient_id,
    )
    result = tuple(row for row in rows if row.snapshot_id in current_ids)

    snapshot_by_id = {snapshot.snapshot_id: snapshot for snapshot in current}
    for row in result:
        snapshot = snapshot_by_id[row.snapshot_id]
        if row.service_date != snapshot.service_date:
            raise PatientSessionHistoryStoreError(
                f"Patient/session row {row.session_id!r} date does not match snapshot {row.snapshot_id!r}"
            )
        if row.snapshot_version != snapshot.version:
            raise PatientSessionHistoryStoreError(
                f"Patient/session row {row.session_id!r} version does not match snapshot {row.snapshot_id!r}"
            )
    return result


def values_for_payload(row: PatientSessionHistoryRow) -> tuple[object | None, ...]:
    """Return the canonical payload order used by close-day/amendment bridges."""

    return (
        row.snapshot_id,
        row.service_date.isoformat(),
        row.snapshot_version,
        row.session_id,
        row.patient_id,
        row.status.value,
        row.original_therapist_id,
        row.original_time.strftime("%H:%M:%S"),
        row.effective_therapist_id,
        row.effective_time.strftime("%H:%M:%S") if row.effective_time is not None else None,
        row.replacement_id,
        row.cancellation_id,
        row.cancellation_kind,
        row.reason,
    )
