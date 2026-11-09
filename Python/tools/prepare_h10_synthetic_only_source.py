from __future__ import annotations

"""One-shot Excel-native isolation of a disposable H8/H10 DEMO history source.

Only a fresh CB candidate is modified. Original CA and every legacy workbook
remain frozen. Never use this tool on patient-care or production workbooks.
"""

import argparse
from hashlib import sha256
from pathlib import Path
import shutil
import sys

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from tools.seed_h8_synthetic_history import HISTORY_SHEETS
from rehab_excel.therapist_attendance_history_store import (
    HEADERS as ATTENDANCE_HEADERS, SHEET_NAME as ATTENDANCE_SHEET,
)
from rehab_excel.day_history_event_store import HEADERS as EVENT_HEADERS
from tools.verify_h8_synthetic_history import verify_demo_history

DEMO_PREFIX = "DEMO-H8-"
# Zero-based logical id columns in the finalized schemas.
# DAY_HISTORY_EVENTS identifies the related snapshot in column K (index 10).
SCHEMAS = {
    name: (tuple(headers), 10 if name == "DAY_HISTORY_EVENTS" else 0)
    for name, headers in HISTORY_SHEETS
}
SCHEMAS[ATTENDANCE_SHEET] = (tuple(ATTENDANCE_HEADERS), 0)


class DemoIsolationError(RuntimeError):
    pass


def sha(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def check_names(source: Path, output: Path) -> None:
    if not source.is_file() or not source.name.upper().endswith("_CA.XLSM") or "_DEMO_" not in source.name.upper():
        raise DemoIsolationError("Existing frozen CA DEMO source required")
    if output.exists() or not output.name.upper().endswith("_CB.XLSM") or "_DEMO_" not in output.name.upper():
        raise DemoIsolationError("Fresh, absent CB DEMO destination required")
    if source.resolve() == output.resolve():
        raise DemoIsolationError("Source/destination collision")


def keep_row(values: tuple[object, ...], *, id_column: int) -> bool:
    """Returns True only for a fully attributed synthetic H8 record."""
    key = values[id_column]
    if key is None or not str(key).strip():
        raise DemoIsolationError("Historical record has no snapshot provenance")
    text = str(key).strip()
    if text.startswith("DEMO-") and not text.startswith(DEMO_PREFIX):
        raise DemoIsolationError("Unknown DEMO namespace in history")
    return text.startswith(DEMO_PREFIX)


def classify_rows(path: Path) -> dict[str, tuple[int, int]]:
    """Read-only strict schema inspection and exact keep/remove counts."""
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        counts = {}
        ids = set()
        for name, (headers, id_col) in SCHEMAS.items():
            if name not in wb:
                raise DemoIsolationError(f"Required history sheet missing: {name}")
            ws = wb[name]
            observed = tuple(
                str(value or "").strip() for value in next(ws.iter_rows(
                    min_row=1, max_row=1, min_col=1, max_col=len(headers),
                    values_only=True,
                ))
            )
            if observed != headers:
                raise DemoIsolationError(f"Invalid historical schema: {name}")
            kept = removed = 0
            for values in ws.iter_rows(
                min_row=2, min_col=1, max_col=len(headers), values_only=True,
            ):
                if not any(value is not None and str(value).strip() for value in values):
                    continue
                if keep_row(values, id_column=id_col):
                    kept += 1
                    if name == "FINALIZED_DAY_SNAPSHOTS":
                        ids.add(str(values[0]).strip())
                else:
                    removed += 1
            counts[name] = (kept, removed)
        if not ids or len(ids) < 12:
            raise DemoIsolationError("Synthetic finalized cohort missing")
        # Cross-sheet strict verification before creating the output.
        verify_demo_history(path)
        return counts
    finally:
        wb.close()


def verify_isolated(path: Path, expected: dict[str, tuple[int, int]]) -> None:
    counts = classify_rows(path)
    for name, (original_kept, _) in expected.items():
        kept, removed = counts[name]
        if kept != original_kept or removed != 0:
            raise DemoIsolationError(
                f"Isolated history mismatch in {name}: kept={kept}, non-demo={removed}"
            )


def install(source: Path, output: Path) -> None:
    check_names(source, output)
    if sys.platform != "win32":
        raise DemoIsolationError("Native Excel on Windows required")
    expected = classify_rows(source)
    original_hash = sha(source)
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise DemoIsolationError("pywin32 required") from exc

    output.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive create; never overwrite an earlier candidate.
    with source.open("rb") as inp, output.open("xb") as out:
        shutil.copyfileobj(inp, out)
    if sha(output) != original_hash or sha(source) != original_hash:
        raise DemoIsolationError("Original-source copy SHA mismatch; freeze CB")

    excel = book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        book = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=False)
        if bool(book.ReadOnly):
            raise DemoIsolationError("Candidate opened read-only")
        for name, (headers, id_col) in SCHEMAS.items():
            ws = book.Worksheets(name)
            if not bool(ws.ProtectContents) or int(ws.Visible) != 2:
                raise DemoIsolationError(f"Technical sheet protection/visibility invalid: {name}")
            # Record row ids and validate before touching this sheet.
            last = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
            to_remove = []
            kept = 0
            for row_idx in range(2, last + 1):
                values = tuple(ws.Cells(row_idx, col).Value for col in range(1, len(headers) + 1))
                if not any(value is not None and str(value).strip() for value in values):
                    continue
                if keep_row(values, id_column=id_col):
                    kept += 1
                else:
                    to_remove.append(row_idx)
            if (kept, len(to_remove)) != expected[name]:
                raise DemoIsolationError(f"Native preflight row count changed: {name}")
            ws.Unprotect()
            # Row deletion from the bottom avoids shifting unrelated ids.
            for row_idx in reversed(to_remove):
                ws.Rows(row_idx).Delete()
            ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
            ws.Visible = 2
        book.Save()
        book.Close(SaveChanges=False)
        book = None
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
    if sha(source) != original_hash:
        raise DemoIsolationError("Frozen CA changed; freeze CB and investigate")
    verify_isolated(output, expected)
    print("H10 SYNTHETIC-ONLY CB SOURCE VERIFIED")
    print("Synthetic finalized history kept; older test fixtures excluded in CB only")
    print("Original CA unchanged: True")
    print("CB remains a disposable non-production demo, not pilot data")


def main() -> int:
    parser = argparse.ArgumentParser(description="Create fresh synthetic-only CB H10 demo")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        install(args.source, args.output)
    except Exception as exc:
        print(f"SAFETY STOP: {type(exc).__name__}: {exc}")
        print("Freeze any created CB candidate; NEVER retry it in place")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
