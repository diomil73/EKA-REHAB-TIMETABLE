"""Capture DAILY_INPUT literal grids and encode a private in-workbook archive.

No workbook writes. This is raw form storage, never an applied absence,
replacement, scheduling approval or attendance record.
"""
from __future__ import annotations

from base64 import b64decode, b64encode
from datetime import date, datetime, time
from hashlib import sha256
import json

from .daily_input_day_switch import DailyInputDayDraft, DailyInputDaySwitchError

ARCHIVE_SHEET = "DAILY_INPUT_DRAFTS"
HEADERS = ("ServiceDate", "Part", "Parts", "PayloadSHA256", "Payload")
CHUNK = 24000


def _area(ws, name):
    matches = []
    for validation in ws.data_validations.dataValidation:
        if validation.type == "list" and str(validation.formula1).lstrip("=").upper() == name:
            matches.extend(validation.sqref.ranges)
    if len(matches) != 1 or matches[0].min_col != 1 or matches[0].max_col != 1:
        raise DailyInputDaySwitchError("Invalid or ambiguous DAILY_INPUT name validation area")
    return matches[0].min_row, matches[0].max_row


def _value(cell):
    if cell.data_type in {"f", "e"}:
        raise DailyInputDaySwitchError("Formula or error in literal DAILY_INPUT data")
    return cell.value


def capture_daily_input(ws, *, parse_date):
    """Use the established Excel date parser and actual named-list row bounds.

    Do not infer row count from UsedRange, populated names, or comments.
    Fail closed on unexpected layout or populated cells beyond editable rows.
    """
    if ws.title != "DAILY_INPUT":
        raise DailyInputDaySwitchError("Expected DAILY_INPUT worksheet")
    t1, t2 = _area(ws, "PY_THERAPISTS")
    p1, p2 = _area(ws, "PY_PATIENTS")
    if t1 != 7 or t2 < t1 or p1 != t2 + 5 or p2 < p1:
        raise DailyInputDaySwitchError("Unexpected DAILY_INPUT row layout")
    for row, headers in (
        (t1 - 1, ("Θεραπευτής", "Όλη ημέρα", "Από", "Έως", "Αιτία", "Σχόλιο")),
        (p1 - 1, ("Ασθενής", "Όλη ημέρα", "Ώρα", "Κατάσταση", "Σχόλιο")),
    ):
        if tuple(_value(ws.cell(row, col)) for col in range(1, len(headers) + 1)) != headers:
            raise DailyInputDaySwitchError("Unexpected DAILY_INPUT headers")
    for row in ws.iter_rows(min_row=t2 + 1, max_row=p1 - 3, max_col=6):
        if any(_value(c) is not None for c in row):
            raise DailyInputDaySwitchError("Unexpected input outside therapist rows")
    if ws.max_row > 10000:
        raise DailyInputDaySwitchError("DAILY_INPUT used area needs explicit layout review")
    if ws.max_row > p2:
        for row in ws.iter_rows(min_row=p2 + 1, max_col=5):
            if any(_value(c) is not None for c in row):
                raise DailyInputDaySwitchError("Unexpected input outside patient rows")
    def grid(first, last, width):
        return tuple(tuple(_value(c) for c in row) for row in
                     ws.iter_rows(min_row=first, max_row=last, max_col=width))
    return DailyInputDayDraft(parse_date(_value(ws["B2"])), grid(t1, t2, 6), grid(p1, p2, 5))


def _encode(value):
    if type(value) in (date, datetime, time):
        return [type(value).__name__, value.isoformat()]
    return ["literal", value]


def _decode(value):
    if not isinstance(value, list) or len(value) != 2:
        raise ValueError("Invalid cell encoding")
    kind, raw = value
    if kind == "literal" and (raw is None or type(raw) in (str, int, float, bool)):
        return raw
    if kind in ("date", "datetime", "time") and type(raw) is str:
        return {"date": date, "datetime": datetime, "time": time}[kind].fromisoformat(raw)
    raise ValueError("Invalid cell encoding")


def archive_rows(drafts):
    """Return literal ASCII payload chunks; safe from Excel formula evaluation.

    Chunking avoids Excel's 32,767-character text-cell limit. The digest is
    corruption detection, not a signature or authorization.
    """
    drafts = tuple(drafts)
    if any(not isinstance(d, DailyInputDayDraft) for d in drafts):
        raise DailyInputDaySwitchError("Invalid archive draft")
    if len({d.service_date for d in drafts}) != len(drafts):
        raise DailyInputDaySwitchError("Duplicate draft date")
    rows = []
    for draft in sorted(drafts, key=lambda d: d.service_date):
        raw = json.dumps([draft.service_date.isoformat(),
                          [[[_encode(v) for v in row] for row in grid]
                           for grid in (draft.therapist_rows, draft.patient_rows)]],
                         ensure_ascii=True, allow_nan=False, separators=(",", ":")).encode("ascii")
        payload = b64encode(raw).decode("ascii")
        chunks = [payload[i:i + CHUNK] for i in range(0, len(payload), CHUNK)]
        digest = sha256(raw).hexdigest()
        for i, chunk in enumerate(chunks, 1):
            rows.append((draft.service_date.isoformat(), i, len(chunks), digest, "b64:" + chunk))
    return tuple(rows)


def decode_archive_rows(rows):
    """Reject malformed, missing, duplicated or altered chunks before returning any draft."""
    try:
        groups = {}
        for day, part, total, digest, payload in rows:
            if (type(day) is not str or date.fromisoformat(day).isoformat() != day
                    or type(part) is not int or type(total) is not int
                    or not 1 <= part <= total
                    or type(digest) is not str or len(digest) != 64
                    or type(payload) is not str or not payload.startswith("b64:")
                    or not 0 < len(payload[4:]) <= CHUNK):
                raise ValueError("Invalid archive row")
            meta, parts = groups.setdefault(day, ((total, digest), {}))
            if meta != (total, digest) or part in parts:
                raise ValueError("Conflicting archive chunks")
            parts[part] = payload[4:]
        result = []
        for day, ((total, digest), parts) in sorted(groups.items()):
            if len(parts) != total:
                raise ValueError("Missing archive chunks")
            raw = b64decode("".join(parts[i] for i in range(1, total + 1)), validate=True)
            if sha256(raw).hexdigest() != digest:
                raise ValueError("Archive digest mismatch")
            record = json.loads(raw)
            if type(record) is not list or len(record) != 2 or record[0] != day:
                raise ValueError("Archive payload date mismatch")
            grids = record[1]
            if type(grids) is not list or len(grids) != 2:
                raise ValueError("Invalid archive grid count")
            if any(type(grid) is not list or any(type(row) is not list for row in grid)
                   for grid in grids):
                raise ValueError("Invalid archive grids")
            decoded = [tuple(tuple(_decode(v) for v in row) for row in grid) for grid in grids]
            result.append(DailyInputDayDraft(date.fromisoformat(day), *decoded))
        return tuple(result)
    except (ValueError, TypeError, KeyError, OverflowError) as exc:
        raise DailyInputDaySwitchError("Invalid or incomplete DAILY_INPUT draft archive") from exc


def read_archive_sheet(ws):
    if ws.title != ARCHIVE_SHEET:
        raise DailyInputDaySwitchError("Unexpected archive worksheet")
    if tuple(_value(ws.cell(1, i)) for i in range(1, 6)) != HEADERS:
        raise DailyInputDaySwitchError("Invalid draft archive headers")
    rows = []
    for cells in ws.iter_rows(min_row=2, max_col=5):
        values = tuple(_value(cell) for cell in cells)
        if any(v is not None for v in values):
            rows.append(values)
    return decode_archive_rows(rows)
