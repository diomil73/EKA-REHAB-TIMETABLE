"""Read-only, source-bound PREPARED handoff for the existing DAILY_INPUT journal.

This preflight DOES NOT modify Excel or stage a durable journal revision.
A later canonical VBA transaction must append the proposed record inside the
open workbook, save and independently verify it BEFORE any operational edit.
"""
from __future__ import annotations

import argparse
from datetime import date, datetime
from hashlib import sha256
import json
from pathlib import Path
import sys
from uuid import uuid4

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_journal import (  # noqa: E402
    JOURNAL_HEADERS, JOURNAL_SHEET, DailyInputJournalError,
    append_operation_journal, decode_operation_journal,
)
from rehab_excel.daily_input_operation import DailyInputOperation  # noqa: E402


class PreparedJournalPreflightError(RuntimeError):
    """A non-mutating pre-write gate failed; no user data in its message."""


def _source_date(value: object) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if type(value) is not str:
        raise PreparedJournalPreflightError("Invalid DAILY_INPUT date")
    for pattern in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(value.strip(), pattern).date().isoformat()
        except ValueError:
            pass
    raise PreparedJournalPreflightError("Invalid DAILY_INPUT date")


def _existing_journal_rows(book) -> tuple[tuple, ...]:
    if JOURNAL_SHEET not in book.sheetnames:
        raise PreparedJournalPreflightError("Existing native journal is required")
    ws = book[JOURNAL_SHEET]
    if ws.sheet_state != "veryHidden" or not bool(ws.protection.sheet):
        raise PreparedJournalPreflightError("Journal protection or visibility invalid")
    if tuple(ws.cell(1, col).value for col in range(1, 6)) != JOURNAL_HEADERS:
        raise PreparedJournalPreflightError("Journal headers invalid")
    # Read all populated cells in the contract columns, then reject any
    # internal blank/partial row instead of silently skipping corruption.
    last = max(
        (r for r in range(1, ws.max_row + 1)
         if any(ws.cell(r, col).value is not None for col in range(1, 6))),
        default=1,
    )
    rows = []
    for row_index in range(2, last + 1):
        values = tuple(ws.cell(row_index, col).value for col in range(1, 6))
        if any(value is None for value in values):
            raise PreparedJournalPreflightError("Partial or missing journal row")
        revision = values[1]
        if type(revision) is float and revision.is_integer():
            values = (values[0], int(revision), *values[2:])
        rows.append(values)
    try:
        decode_operation_journal(rows)
    except DailyInputJournalError:
        raise PreparedJournalPreflightError("Existing journal history invalid") from None
    return tuple(rows)


def journal_prefix_sha256(rows: tuple[tuple, ...]) -> str:
    """Detect accidental changes to any earlier journal row, not authenticity."""
    canonical = json.dumps(
        [list(row) for row in rows],
        ensure_ascii=True, allow_nan=False, separators=(",", ":"),
    )
    return sha256(("diomil-journal-history-v1:" + canonical).encode("ascii")).hexdigest()


def prepare_prepared_handoff(
    source_path: str | Path,
    preparation_response_path: str | Path,
    payload_path: str | Path,
    *,
    operation_id: str | None = None,
) -> dict[str, object]:
    """Produce canonical PREPARED row metadata with no source or Excel writes.

    The response must be the current successful output of the already-canonical
    prepare_daily_input_in_place tool for the same unmodified source and payload.
    """
    source = Path(source_path).resolve()
    response = Path(preparation_response_path).resolve()
    payload = Path(payload_path).resolve()
    if source.suffix.casefold() != ".xlsm" or not source.is_file():
        raise PreparedJournalPreflightError("Expected existing .xlsm source")
    if not response.is_file() or not payload.is_file():
        raise PreparedJournalPreflightError("Preparation evidence is missing")
    if source in (response, payload) or response == payload:
        raise PreparedJournalPreflightError("Preflight file paths overlap")

    try:
        receipt = json.loads(response.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, ValueError):
        raise PreparedJournalPreflightError("Invalid preparation receipt") from None
    if (type(receipt) is not dict or receipt.get("ok") is not True
            or receipt.get("source_unchanged") is not True
            or type(receipt.get("date")) is not str
            or type(receipt.get("source_sha256_before")) is not str
            or type(receipt.get("payload_path")) is not str):
        raise PreparedJournalPreflightError("Preparation receipt not verified")
    if Path(receipt["payload_path"]).resolve() != payload:
        raise PreparedJournalPreflightError("Prepared payload path mismatch")
    source_hash = file_sha256(source)
    if source_hash != receipt["source_sha256_before"]:
        raise PreparedJournalPreflightError("Workbook changed after preparation")
    try:
        target_date = date.fromisoformat(receipt["date"]).isoformat()
    except ValueError:
        raise PreparedJournalPreflightError("Invalid preparation service date") from None

    book = None
    try:
        book = load_workbook(source, read_only=False, data_only=False, keep_vba=True)
        if "DAILY_INPUT" not in book.sheetnames:
            raise PreparedJournalPreflightError("Missing DAILY_INPUT")
        workbook_date = _source_date(book["DAILY_INPUT"]["B2"].value)
        if workbook_date != target_date:
            raise PreparedJournalPreflightError("Preparation service date mismatch")
        before = _existing_journal_rows(book)
    except PreparedJournalPreflightError:
        raise
    except Exception:
        raise PreparedJournalPreflightError("Source workbook preflight failed") from None
    finally:
        if book is not None:
            book.close()

    if file_sha256(source) != source_hash:
        raise PreparedJournalPreflightError("Workbook changed during preflight")
    payload_hash = file_sha256(payload)
    if not payload_hash:
        raise PreparedJournalPreflightError("Payload hash unavailable")
    new_id = operation_id if operation_id is not None else "diomil-daily-" + uuid4().hex
    try:
        prepared = DailyInputOperation(
            operation_id=new_id, source_sha256=source_hash, service_date=target_date,
        )
        expected = append_operation_journal(before, prepared)
    except (ValueError, DailyInputJournalError):
        raise PreparedJournalPreflightError("New journal operation rejected") from None
    if len(expected) != len(before) + 1 or expected[:-1] != before:
        raise PreparedJournalPreflightError("Prepared journal append is not new")
    if file_sha256(source) != source_hash or file_sha256(payload) != payload_hash:
        raise PreparedJournalPreflightError("Source or payload changed during preflight")

    return {
        "status": "PREPARED_NOT_PERSISTED",
        "operation_id": new_id,
        "service_date": target_date,
        "source_sha256": source_hash,
        "payload_sha256": payload_hash,
        "existing_journal_rows": len(before),
        "journal_prefix_sha256": journal_prefix_sha256(before),
        "prepared_row": list(expected[-1]),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only DAILY_INPUT PREPARED journal handoff.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--preparation-response", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    created_output = False
    output = args.output.resolve()
    try:
        if output in {args.input.resolve(), args.preparation_response.resolve(), args.payload.resolve()}:
            raise PreparedJournalPreflightError("Output overlaps preflight input")
        if output.exists():
            raise PreparedJournalPreflightError("Handoff output already exists")
        record = prepare_prepared_handoff(
            args.input, args.preparation_response, args.payload,
        )
        # Exclusive create: never replace a prior proposed operation record.
        with output.open("x", encoding="utf-8") as handle:
            created_output = True
            json.dump(record, handle, ensure_ascii=True, sort_keys=True, separators=(",", ":"))
            handle.write("\n")
        if (file_sha256(args.input) != record["source_sha256"]
                or file_sha256(args.payload) != record["payload_sha256"]):
            raise PreparedJournalPreflightError("Source or payload changed before handoff completion")
    except Exception:
        # Clean up only output that this invocation created, not an older handoff.
        if created_output:
            try:
                output.unlink(missing_ok=True)
            except OSError:
                pass
        # No data-bearing exception messages or paths on stdout/stderr.
        print("SAFETY STOP: DAILY_INPUT PREPARED PREFLIGHT FAILED")
        return 2
    print("DAILY_INPUT PREPARED HANDOFF READY: NOT PERSISTED")
    print("Workbook unchanged: YES")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
