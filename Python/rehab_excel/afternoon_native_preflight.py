"""Read-only structural preflight for a disposable native patient-panel XLSM.

No COM, no workbook save, no patient IDs/names in diagnostics. This is only
structural eligibility; it cannot establish synthetic provenance or approve a
native installer. The operator must separately confirm source provenance.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

REQUIRED_SHEETS = frozenset({"PATIENTS", "PATIENT_PLANNER", "MASTER_SCHEDULE"})
AFTERNOON_SHEET = "AFTERNOON_PHYSIO_SLOTS"
EXPECTED_HEADERS = ("PatientID", "Treatment", "Weekday", "StartTime", "Therapist")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


class AfternoonNativePreflightError(ValueError):
    """XLSM structure cannot safely be admitted as a disposable preview."""


@dataclass(frozen=True)
class AfternoonNativePreflight:
    source_name: str
    sheet_count: int
    has_afternoon_sheet: bool
    has_vba_binary: bool
    needs_sheet_install: bool
    ready_for_provenance_review: bool


def inspect_afternoon_native_candidate(path: str | Path) -> AfternoonNativePreflight:
    """Inspect ZIP/XML structure without opening in Excel or reading patient rows.

    A workbook with an existing afternoon sheet is explicitly *not* eligible
    for first-time sheet installation. This does not validate sheet headers,
    existing VBA forms or macro behavior, which require separate native proof.
    """
    candidate = Path(path)
    if (candidate.suffix.casefold() != ".xlsm"
            or not candidate.is_file() or candidate.is_symlink()):
        raise AfternoonNativePreflightError("Existing regular XLSM required")
    try:
        with ZipFile(candidate, "r") as archive:
            members = set(archive.namelist())
            if "xl/workbook.xml" not in members or "xl/vbaProject.bin" not in members:
                raise AfternoonNativePreflightError("XLSM workbook or VBA binary missing")
            root = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets = root.findall("m:sheets/m:sheet", namespaces=NS)
            names = [sheet.attrib.get("name", "") for sheet in sheets]
            if not names or len(names) != len(set(names)):
                raise AfternoonNativePreflightError("Invalid worksheet roster")
            missing = REQUIRED_SHEETS - set(names)
            if missing:
                raise AfternoonNativePreflightError(
                    "Required application worksheet missing"
                )
            if AFTERNOON_SHEET in names:
                raise AfternoonNativePreflightError(
                    "Existing afternoon sheet: first-time installer must not run"
                )
            return AfternoonNativePreflight(
                source_name=candidate.name,
                sheet_count=len(names),
                has_afternoon_sheet=False,
                has_vba_binary=True,
                needs_sheet_install=True,
                ready_for_provenance_review=True,
            )
    except (BadZipFile, ET.ParseError, OSError, KeyError) as exc:
        raise AfternoonNativePreflightError("Invalid or unreadable XLSM structure") from exc
