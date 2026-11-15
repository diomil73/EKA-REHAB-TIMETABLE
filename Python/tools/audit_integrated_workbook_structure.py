from __future__ import annotations

"""Non-identifying, read-only structural inventory of one Excel workbook.

This is an architecture/provenance preflight only. It never prints cell values,
patient/staff names, identifiers, clinical notes, formulas, or cached values.
It never saves or rewrites the source workbook.
"""

import argparse
from hashlib import sha256
from pathlib import Path
import re
import sys
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook


ALLOWED_SUFFIXES = {".xlsm", ".xlsx"}
MAX_CELLS_PER_SHEET = 1_000_000
MAX_TOTAL_CELLS = 3_000_000
HEADER_SCAN_ROWS = 30

ROLE_PATTERNS = {
    "patient": re.compile(r"patient|ασθεν", re.IGNORECASE),
    "therapist": re.compile(r"therapist|θεραπευτ", re.IGNORECASE),
    "doctor": re.compile(r"doctor|ιατρ|γιατρ", re.IGNORECASE),
    "student": re.compile(r"student|σπουδαστ|φοιτητ", re.IGNORECASE),
    "assignment": re.compile(r"assignment|ανάθεσ|αναθεσ", re.IGNORECASE),
    "session": re.compile(r"session|συνεδρ|θεραπ", re.IGNORECASE),
    "attendance": re.compile(r"attendance|παρουσ|απουσ", re.IGNORECASE),
    "date_time": re.compile(r"date|time|ημερ|ημέρ|ωρα|ώρα", re.IGNORECASE),
    "room": re.compile(r"room|ward|θαλαμ|κλινικ", re.IGNORECASE),
    "afternoon": re.compile(r"afternoon|απογευμα", re.IGNORECASE),
    "pool": re.compile(r"pool|πισιν", re.IGNORECASE),
    "robotics": re.compile(r"robot|ρομποτ", re.IGNORECASE),
    "psychology": re.compile(r"psych|ψυχ", re.IGNORECASE),
    "physiotherapy": re.compile(r"physio|φυσικοθεραπ|φ/θ", re.IGNORECASE),
}

ZIP_GROUPS = {
    "vba_project": lambda n: n.casefold() == "xl/vbaproject.bin",
    "external_links": lambda n: n.casefold().startswith("xl/externallinks/"),
    "connections": lambda n: n.casefold() == "xl/connections.xml",
    "query_tables": lambda n: n.casefold().startswith("xl/querytables/"),
    "pivot_tables": lambda n: n.casefold().startswith("xl/pivottables/"),
    "pivot_caches": lambda n: n.casefold().startswith("xl/pivotcache/"),
    "slicers": lambda n: "slicer" in n.casefold(),
    "drawings": lambda n: n.casefold().startswith("xl/drawings/"),
    "charts": lambda n: n.casefold().startswith("xl/charts/"),
    "activex": lambda n: n.casefold().startswith("xl/activex/"),
    "embeddings": lambda n: n.casefold().startswith("xl/embeddings/"),
    "custom_ui": lambda n: n.casefold().startswith("customui/"),
    "comments": lambda n: "comment" in n.casefold(),
    "vml_drawings": lambda n: n.casefold().startswith("xl/drawings/vml"),
    "external_workbook_links": lambda n: n.casefold().startswith("xl/externalLinks/".casefold()),
}


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _roles(value: object) -> set[str]:
    if not isinstance(value, str) or not value.strip():
        return set()
    return {role for role, pattern in ROLE_PATTERNS.items() if pattern.search(value)}


def _zip_inventory(path: Path) -> dict[str, object]:
    with ZipFile(path, "r") as archive:
        names = archive.namelist()
        corrupt = archive.testzip()
        counts = {
            group: sum(1 for name in names if predicate(name))
            for group, predicate in ZIP_GROUPS.items()
        }
        vba = next(
            (info for info in archive.infolist()
             if info.filename.casefold() == "xl/vbaproject.bin"),
            None,
        )
        return {
            "zip_entries": len(names),
            "zip_corrupt_member": corrupt,
            "zip_groups": counts,
            "vba_project_bytes": 0 if vba is None else vba.file_size,
            "vba_project_sha256": None if vba is None else sha256(archive.read(vba)).hexdigest(),
            "workbook_xml_present": "xl/workbook.xml" in names,
        }


def audit(path: Path) -> int:
    path = path.expanduser().resolve()
    if not path.is_file() or path.suffix.casefold() not in ALLOWED_SUFFIXES:
        print("SAFETY STOP: Existing XLSM or XLSX file required.")
        return 2

    before_hash = _sha256(path)
    before_size = path.stat().st_size
    formulas_wb = None
    cached_wb = None
    incomplete: list[str] = []
    total_cells = 0

    try:
        zip_info = _zip_inventory(path)
        if zip_info["zip_corrupt_member"] is not None:
            print("SAFETY STOP: XLSX container has a corrupt ZIP member.")
            return 2
        if not zip_info["workbook_xml_present"]:
            print("SAFETY STOP: Excel workbook manifest is missing.")
            return 2

        formulas_wb = load_workbook(
            path, read_only=True, data_only=False, keep_vba=False, keep_links=False
        )
        cached_wb = load_workbook(
            path, read_only=True, data_only=True, keep_vba=False, keep_links=False
        )

        print("DIOMIL READ-ONLY INTEGRATED WORKBOOK STRUCTURE AUDIT")
        print("Candidate filename:", path.name)
        print("File extension:", path.suffix.lower())
        print("File size bytes:", before_size)
        print("File SHA256:", before_hash)
        print("ZIP entries:", zip_info["zip_entries"])
        print("ZIP groups:", ", ".join(
            f"{key}={value}" for key, value in zip_info["zip_groups"].items()
        ))
        print("VBA project bytes:", zip_info["vba_project_bytes"])
        if zip_info["vba_project_sha256"]:
            print("VBA project SHA256:", zip_info["vba_project_sha256"])
        print("Defined names:", len(formulas_wb.defined_names))
        print("Worksheet count:", len(formulas_wb.worksheets))
        print("CELL VALUES, FORMULAS AND IDENTITIES ARE NOT PRINTED")
        print("")

        patient_staff_indices: list[int] = []
        for index, ws in enumerate(formulas_wb.worksheets, start=1):
            cached_ws = cached_wb[ws.title]
            dimension = ws.calculate_dimension()
            sheet_cells = 0
            nonempty = 0
            formula_count = 0
            cached_formula_count = 0
            header_roles: set[str] = set()
            sheet_limit_hit = False

            formula_rows = ws.iter_rows()
            cached_rows = cached_ws.iter_rows()
            for row_number, (formula_row, cached_row) in enumerate(
                zip(formula_rows, cached_rows), start=1
            ):
                for formula_cell, cached_cell in zip(formula_row, cached_row):
                    sheet_cells += 1
                    total_cells += 1
                    value = formula_cell.value
                    if value is not None:
                        nonempty += 1
                    if formula_cell.data_type == "f":
                        formula_count += 1
                        if cached_cell.value is not None:
                            cached_formula_count += 1
                    if row_number <= HEADER_SCAN_ROWS:
                        header_roles.update(_roles(value))
                    if sheet_cells >= MAX_CELLS_PER_SHEET or total_cells >= MAX_TOTAL_CELLS:
                        sheet_limit_hit = True
                        break
                if sheet_limit_hit:
                    break

            name_roles = _roles(ws.title)
            combined_roles = header_roles | name_roles
            is_patient_staff = bool(
                combined_roles.intersection(
                    {"patient", "therapist", "doctor", "student", "assignment",
                     "session", "attendance", "physiotherapy", "psychology"}
                )
            )
            if is_patient_staff:
                patient_staff_indices.append(index)

            role_label = ",".join(sorted(combined_roles)) or "unclassified"
            print(
                f"SHEET {index:03d} | name={ws.title!r} | state={ws.sheet_state} | "
                f"dimension={dimension} | scanned_cells={sheet_cells} | "
                f"nonempty_cells={nonempty} | formulas={formula_count} | "
                f"formula_cached_results={cached_formula_count} | "
                f"role_hints={role_label} | patient_staff_surface={'YES' if is_patient_staff else 'NO'}"
            )

            if sheet_limit_hit:
                incomplete.append(ws.title)
                print("  SCAN LIMIT REACHED; sheet is not fully inventoried.")
                if total_cells >= MAX_TOTAL_CELLS:
                    incomplete.append("TOTAL_CELL_LIMIT")
                    break

        print("")
        print("PATIENT_OR_STAFF_BEARING_SHEET_INDICES:",
              ",".join(str(x) for x in patient_staff_indices) or "NONE_DETECTED")
        print("TOTAL_SCANNED_CELLS:", total_cells)
        print("STRUCTURAL_SCAN_LIMITS:",
              f"per_sheet={MAX_CELLS_PER_SHEET}; total={MAX_TOTAL_CELLS}")
        if incomplete:
            print("STRUCTURE AUDIT INCOMPLETE:", ",".join(incomplete))

    except (BadZipFile, OSError, ValueError, KeyError, TypeError) as exc:
        print(f"SAFETY STOP: Read-only structure inspection failed ({type(exc).__name__}).")
        return 2
    except Exception as exc:
        print(f"SAFETY STOP: Unexpected read-only inspection failure ({type(exc).__name__}).")
        return 2
    finally:
        if cached_wb is not None:
            cached_wb.close()
        if formulas_wb is not None:
            formulas_wb.close()

    after_hash = _sha256(path)
    after_size = path.stat().st_size
    if (after_hash, after_size) != (before_hash, before_size):
        print("SAFETY STOP: Source workbook changed during read-only audit.")
        return 2

    print("SOURCE HASH UNCHANGED: YES")
    if incomplete:
        print("READ-ONLY STRUCTURE AUDIT INCOMPLETE")
        return 3
    print("READ-ONLY STRUCTURE AUDIT COMPLETE")
    print("NO WORKBOOK WAS MODIFIED")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Non-identifying, read-only inventory of one integrated Excel workbook."
    )
    parser.add_argument(
        "workbooks",
        nargs="+",
        type=Path,
        help="One or more existing XLSM/XLSX candidates to inventory read-only.",
    )
    args = parser.parse_args(argv)
    results = []
    for candidate in args.workbooks:
        print("\\n=== CANDIDATE START ===")
        results.append(audit(candidate))
        print("=== CANDIDATE END ===")
    if all(code == 0 for code in results):
        print("ALL CANDIDATE STRUCTURE AUDITS COMPLETE")
        return 0
    print("CANDIDATE STRUCTURE AUDITS REQUIRE REVIEW:", results)
    return 3 if 3 in results else 2


if __name__ == "__main__":
    raise SystemExit(main())
