"""Strict read-only saved-OOXML MASTER_SCHEDULE projection diagnostics.

This is a deliberately scoped comparison of the *structural* MASTER
projection against independently supplied expected MasterProjectionRow
facts. It does not verify evaluated Excel formula results, therapist
allocation, daily scheduling, native Excel behavior or phase authorization.
The source workbook must not be used to invent the expected projection.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path

from openpyxl import load_workbook

from .master_projection import MasterProjectionRow

SOURCE_COLUMNS = {
    1: "D", 2: "B", 3: "C", 4: "F", 5: "I",
    6: "N", 7: "L", 8: "P", 9: "R", 10: "T",
}


class MasterProjectionSavedReadbackError(RuntimeError):
    """Saved MASTER structure fails strict independent readback."""


@dataclass(frozen=True)
class MasterProjectionSavedReadback:
    workbook_sha256: str
    projected_rows: int
    header_row: int
    status_column: int


def _digest(path: Path) -> str:
    h = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_master_projection_saved_readback(
    workbook_path: str | Path,
    *,
    expected_projection: tuple[MasterProjectionRow, ...],
) -> MasterProjectionSavedReadback:
    """Fail closed on formula, row order, identity, schema or tail mismatch.

    The expectation is obtained from independent preflight/source facts, not
    derived from the reopened MASTER itself. Checks formula literals rather
    than Excel-calculated results. No workbook writes or macros are performed.
    """
    if type(expected_projection) is not tuple:
        raise MasterProjectionSavedReadbackError("Explicit immutable expected projection required")
    if any(type(x) is not MasterProjectionRow for x in expected_projection):
        raise MasterProjectionSavedReadbackError("Expected projection contains invalid entries")
    rows = [x.target_row for x in expected_projection]
    planner_rows = [x.planner_row for x in expected_projection]
    ids = [x.patient_id for x in expected_projection]
    if (len(rows) != len(set(rows)) or rows != sorted(rows)
            or any(type(r) is not int or r < 2 for r in rows)
            or any(type(r) is not int or r < 2 for r in planner_rows)
            or len(ids) != len(set(ids))
            or any(type(i) is not str or not i.strip() for i in ids)):
        raise MasterProjectionSavedReadbackError("Invalid expected projection identity or order")

    try:
        path = Path(workbook_path).resolve(strict=True)
        if not path.is_file() or path.suffix.casefold() != ".xlsm":
            raise MasterProjectionSavedReadbackError("Existing saved XLSM required")
        before = _digest(path)
        wb = load_workbook(path, read_only=False, data_only=False, keep_vba=False)
    except MasterProjectionSavedReadbackError:
        raise
    except Exception as exc:
        raise MasterProjectionSavedReadbackError("Saved MASTER workbook cannot be opened") from exc

    try:
        if "MASTER_SCHEDULE" not in wb.sheetnames:
            raise MasterProjectionSavedReadbackError("MASTER_SCHEDULE missing")
        ws = wb["MASTER_SCHEDULE"]
        header_rows = [
            row for row in range(1, 11)
            if str(ws.cell(row, 2).value or "").strip().casefold() == "Θάλαμος".casefold()
            and str(ws.cell(row, 3).value or "").strip().casefold() == "Ασθενής".casefold()
        ]
        if len(header_rows) != 1:
            raise MasterProjectionSavedReadbackError("MASTER header missing or ambiguous")
        header = header_rows[0]
        status_header = str(ws.cell(header, 13).value or "").strip().casefold()
        if status_header == "κατάσταση":
            if (str(ws.cell(header, 11).value or "").strip().casefold() != "ψυχολογοι"
                    or str(ws.cell(header, 12).value or "").strip().casefold() != "απογευματινο προγραμμα"):
                raise MasterProjectionSavedReadbackError("Expanded MASTER schema invalid")
            status = 13
        elif str(ws.cell(header, 11).value or "").strip().casefold() == "κατάσταση":
            status = 11
        else:
            raise MasterProjectionSavedReadbackError("Unsupported MASTER status column")

        expanded_schema = status == 13
        if expanded_schema and any(
            type(item.psychology_display) is not str
            for item in expected_projection
        ):
            raise MasterProjectionSavedReadbackError(
                "Expanded MASTER requires independent psychology expectations"
            )
        if not expanded_schema and any(
            item.psychology_display is not None
            for item in expected_projection
        ):
            raise MasterProjectionSavedReadbackError(
                "Legacy MASTER cannot accept expanded psychology expectations"
            )

        if rows and rows[0] <= header:
            raise MasterProjectionSavedReadbackError("Projection overlaps header")
        expected_by_row = {item.target_row: item for item in expected_projection}
        last = max(rows, default=header)
        # The writer can reserve blank clinic separators. Reject undeclared
        # business rows in those gaps, while preserving decorative formatting.
        for row in range(header + 1, last + 2):
            item = expected_by_row.get(row)
            if item is None:
                for col in range(1, status + 1):
                    if ws.cell(row, col).value not in (None, ""):
                        raise MasterProjectionSavedReadbackError("Unexpected MASTER row or tail data")
                continue
            formulas = {
                col: f"=PATIENT_PLANNER!{source}{item.planner_row}"
                for col, source in SOURCE_COLUMNS.items()
            }
            formulas[status] = f"=PATIENT_PLANNER!E{item.planner_row}"
            if item.responsible_doctor:
                expected_identity = f"{item.display_name}\n✚ {item.responsible_doctor}"
                if ws.cell(row, 3).value != expected_identity:
                    raise MasterProjectionSavedReadbackError("Doctor/patient identity mismatch")
                formulas.pop(3)
            for col, expected in formulas.items():
                cell = ws.cell(row, col)
                if cell.data_type != "f" or cell.value != expected:
                    raise MasterProjectionSavedReadbackError("MASTER formula linkage mismatch")
            if expanded_schema:
                psychology_cell = ws.cell(row, 11)
                if psychology_cell.data_type == "f" or psychology_cell.value != item.psychology_display:
                    raise MasterProjectionSavedReadbackError(
                        "MASTER psychology literal differs from independent preflight"
                    )
                if ws.cell(row, 12).value not in (None, ""):
                    raise MasterProjectionSavedReadbackError(
                        "MASTER afternoon column must remain blank until approved"
                    )
            for col in (14, 15, 16):
                if ws.cell(row, col).value not in (None, ""):
                    raise MasterProjectionSavedReadbackError("Unexpected MASTER legacy tail data")
        # For the 13-column template, K/L are deliberately outside canonical
        # source formulas. Their *evaluated* psychology/afternoon values and
        # therapist views need their own independent phase verifier.
        if _digest(path) != before:
            raise MasterProjectionSavedReadbackError("Workbook changed during MASTER readback")
        return MasterProjectionSavedReadback(before, len(expected_projection), header, status)
    except MasterProjectionSavedReadbackError:
        raise
    except Exception as exc:
        raise MasterProjectionSavedReadbackError("Saved MASTER projection readback invalid") from exc
    finally:
        wb.close()
