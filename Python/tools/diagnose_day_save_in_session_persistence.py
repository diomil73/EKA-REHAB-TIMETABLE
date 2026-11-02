from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from tools.prepare_day_amendment_in_session import prepare_day_amendment_in_session  # noqa: E402


def _parse_payload(path: Path):
    records: list[tuple[str, list[str]]] = []
    text = path.read_text(encoding="utf-8-sig")
    for raw in text.splitlines():
        if not raw.strip():
            continue
        parts = raw.split("\t")
        records.append((parts[0], parts[1:]))
    return records


def _cell_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _row_text(ws, row: int, count: int) -> list[str]:
    return [_cell_text(ws.Cells(row, col).Value) for col in range(1, count + 1)]


def _next_row(ws) -> int:
    xl_up = -4162
    row = int(ws.Cells(ws.Rows.Count, 1).End(xl_up).Row)
    return max(row + 1, 2)


def _find_id(ws, record_id: str) -> int:
    xl_up = -4162
    last = int(ws.Cells(ws.Rows.Count, 1).End(xl_up).Row)
    for row in range(2, last + 1):
        if _cell_text(ws.Cells(row, 1).Value) == record_id:
            return row
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Diagnose in-session day-save persistence on a disposable workbook copy.")
    parser.add_argument("workbook", type=Path)
    args = parser.parse_args()

    source = args.workbook.resolve()
    if not source.exists():
        print(f"DIAGNOSTIC FAILED: workbook not found: {source}")
        return 2

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("DIAGNOSTIC FAILED: pywin32 is required")
        return 2

    with tempfile.TemporaryDirectory(prefix="eka_day_save_diag_") as temp_dir:
        temp = Path(temp_dir)
        copy_path = temp / source.name
        payload = temp / "payload.tsv"
        shutil.copy2(source, copy_path)
        prepare_day_amendment_in_session(copy_path, payload, actor_or_source="diag:day-save")
        records = _parse_payload(payload)

        supersede = next(values for kind, values in records if kind == "SUPERSEDE")
        snapshot = next(values for kind, values in records if kind == "SNAPSHOT")
        therapists = [values for kind, values in records if kind == "THERAPIST"]
        events = [values for kind, values in records if kind == "EVENT"]

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(str(copy_path), UpdateLinks=0, ReadOnly=False)

            ws_snap = workbook.Worksheets("FINALIZED_DAY_SNAPSHOTS")
            ws_th = workbook.Worksheets("THERAPIST_DAY_HISTORY")
            ws_ev = workbook.Worksheets("DAY_HISTORY_EVENTS")
            ws_snap.Unprotect()
            ws_th.Unprotect()
            ws_ev.Unprotect()

            old_row = _find_id(ws_snap, supersede[0])
            if old_row == 0:
                print("FAIL stage=preflight error=previous_snapshot_not_found")
                return 1
            if _cell_text(ws_snap.Cells(old_row, 4).Value).upper() != "CURRENT_FINAL":
                print(f"FAIL stage=preflight error=previous_status actual={_cell_text(ws_snap.Cells(old_row, 4).Value)!r}")
                return 1

            print(f"PASS stage=preflight previous_row={old_row}")
            old_status = ws_snap.Cells(old_row, 4).Value
            ws_snap.Cells(old_row, 4).Value2 = supersede[3]
            actual_identity = [_cell_text(ws_snap.Cells(old_row, 1).Value), _cell_text(ws_snap.Cells(old_row, 3).Value), _cell_text(ws_snap.Cells(old_row, 4).Value).upper()]
            expected_identity = [supersede[0], supersede[2], supersede[3].upper()]
            if actual_identity != expected_identity:
                print(f"FAIL stage=supersede_readback expected={expected_identity!r} actual={actual_identity!r}")
                return 1
            print("PASS stage=supersede_readback")

            new_row = _next_row(ws_snap)
            for col, value in enumerate(snapshot, 1):
                ws_snap.Cells(new_row, col).Value2 = value
            actual = _row_text(ws_snap, new_row, len(snapshot))
            if actual != snapshot:
                for idx, (exp, got) in enumerate(zip(snapshot, actual), 1):
                    if exp != got:
                        print(f"FAIL stage=replacement_snapshot_readback col={idx} expected={exp!r} actual={got!r}")
                        break
                return 1
            print("PASS stage=replacement_snapshot_readback")

            th_row = _next_row(ws_th)
            for offset, values in enumerate(therapists):
                target = th_row + offset
                for col, value in enumerate(values, 1):
                    ws_th.Cells(target, col).Value2 = value
                actual = _row_text(ws_th, target, len(values))
                if actual != values:
                    for idx, (exp, got) in enumerate(zip(values, actual), 1):
                        if exp != got:
                            print(f"FAIL stage=therapist_readback row={offset + 1} col={idx} therapist={values[2]!r} expected={exp!r} actual={got!r}")
                            break
                    return 1
            print(f"PASS stage=therapist_readback rows={len(therapists)}")

            ev_row = _next_row(ws_ev)
            for offset, values in enumerate(events):
                target = ev_row + offset
                for col, value in enumerate(values, 1):
                    ws_ev.Cells(target, col).Value2 = value
                actual = _row_text(ws_ev, target, len(values))
                if actual != values:
                    for idx, (exp, got) in enumerate(zip(values, actual), 1):
                        if exp != got:
                            print(f"FAIL stage=event_readback row={offset + 1} col={idx} expected={exp!r} actual={got!r}")
                            break
                    return 1
            print(f"PASS stage=event_readback rows={len(events)}")

            ws_snap.Cells(old_row, 4).Value2 = old_status
            print("DAY SAVE PERSISTENCE DIAGNOSTIC OK")
            return 0
        except Exception as exc:
            print(f"DIAGNOSTIC FAILED: {type(exc).__name__}: {exc}")
            return 2
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


if __name__ == "__main__":
    raise SystemExit(main())
