"""Read-only, conservative column-C parser for irregular Greek manual schedules.

Outputs PROPOSALS only. No appointment writes, no history writes, no patient identifiers.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

# Provisional whitelist: expand only via independently reviewed therapist aliases.
KNOWN_THERAPISTS = ("ΨΙΑΧΟΥ", "ΚΑΛΥΒΙΩΤΗΣ", "ΓΑΒΡΑΣ", "ΖΗΜΑΡΗΣ", "ΣΑΚΚΟΥ", "ΜΗΛΙΔΑΚΗΣ", "ΧΡΗΣΤΟΥ")
TIME_PATTERN = re.compile(
    r"(?<![0-9])(?:[01]?\d|2[0-3])\s*(?:\.{1,2}|:|,)\s*[0-5]\d(?!\d)"
    r"|(?<![0-9])(?:[01]?\d|2[0-3])[0-5]\d(?!\d)"
    r"|(?<![0-9])[0-9][0-5]\d(?!\d)"
)
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


@dataclass(frozen=True)
class Interpretation:
    status: str
    initial_time: str = ""
    initial_therapist: str = ""
    replacement_time: str = ""
    replacement_therapist: str = ""
    note: str = ""


def _format_time(value: str) -> str | None:
    compact = re.sub(r"\s|\.{1,2}|:|,", "", value)
    if not compact.isdigit() or len(compact) not in (3, 4):
        return None
    h, m = int(compact[:-2]), int(compact[-2:])
    if not (0 <= h <= 23 and 0 <= m <= 59):
        return None
    return f"{h:02d}:{m:02d}"


def parse_physiotherapy_cell(text: str, therapists: tuple[str, ...] = KNOWN_THERAPISTS) -> Interpretation:
    """Propose an initial appointment or a linked replacement, never approve either."""
    original = str(text or "").strip()
    if not original:
        return Interpretation("EMPTY")
    normalized = original.upper().replace("\u00a0", " ").replace("\u200b", "")
    events = []
    occupied = []
    for match in TIME_PATTERN.finditer(normalized):
        parsed = _format_time(match.group())
        if parsed:
            events.append((match.start(), match.end(), "time", parsed))
            occupied.append(match.span())
    for name in sorted(set(therapists), key=len, reverse=True):
        expression = re.compile(
            r"(?<![Α-ΩΆΈΉΊΌΎΏΪΫ])" + re.escape(name)
            + r"(?![Α-ΩΆΈΉΊΌΎΏΪΫ])"
        )
        for match in expression.finditer(normalized):
            if not any(start < match.end() and match.start() < end for start, end in occupied):
                events.append((match.start(), match.end(), "therapist", name))
                occupied.append(match.span())
    events.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    times = [e for e in events if e[2] == "time"]
    names = [e for e in events if e[2] == "therapist"]
    if len(times) > 2 or len(names) > 2:
        return Interpretation("REVIEW", note="Multiple time/name candidates")
    if not times or not names:
        return Interpretation("REVIEW", note="Initial time or recognized therapist absent")
    if times[0][0] > names[0][0]:
        return Interpretation("REVIEW", note="Unrecognized initial time/name ordering")
    initial_time, initial_therapist = times[0][3], names[0][3]
    if len(times) == 1 and len(names) == 1:
        return Interpretation("PROPOSED_INITIAL", initial_time, initial_therapist)
    if len(names) == 2:
        extra_name = names[1]
        if extra_name[0] <= names[0][0]:
            return Interpretation("REVIEW", note="Unclear therapist sequence")
        if len(times) == 2:
            if times[1][0] < names[0][1] or times[1][0] > extra_name[0]:
                return Interpretation("REVIEW", note="Second time not tied clearly to replacement")
            effective_time = times[1][3]
        else:
            effective_time = initial_time
        return Interpretation(
            "PROPOSED_REPLACEMENT", initial_time, initial_therapist,
            effective_time, extra_name[3],
            "Not a second independent appointment"
        )
    if len(times) == 2 and times[1][0] > names[0][1]:
        return Interpretation(
            "REVIEW", initial_time, initial_therapist,
            times[1][3], initial_therapist,
            "Time-only change requires confirmation"
        )
    return Interpretation("REVIEW", initial_time, initial_therapist,
                          note="Ambiguous additional tokens")


def _visible_cells(path: Path, column: str = "C") -> list[tuple[str, str]]:
    """Read literal source cells directly; do not save/recalculate the source XLSX."""
    with ZipFile(path) as z:
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        relationships = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"].lstrip("/") for r in relationships}
        sheet = next(
            (s for s in workbook.findall("m:sheets/m:sheet", NS)
             if s.attrib.get("name") == "ΠΡΟΓΡΑΜΜΑ"),
            None
        )
        if sheet is None or sheet.attrib.get("state", "visible") != "visible":
            raise ValueError("Required visible ΠΡΟΓΡΑΜΜΑ sheet is missing")
        target = targets[sheet.attrib[f"{{{NS['r']}}}id"]]
        if not target.startswith("xl/"):
            target = "xl/" + target
        root = ET.fromstring(z.read(target))
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            shared = [
                "".join(s.itertext())
                for s in ET.fromstring(z.read("xl/sharedStrings.xml"))
            ]
        result = []
        for cell in root.findall(".//m:sheetData/m:row/m:c", NS):
            address = cell.attrib.get("r", "")
            if not re.fullmatch(re.escape(column) + r"\d+", address):
                continue
            v, inline = cell.find("m:v", NS), cell.find("m:is", NS)
            if cell.attrib.get("t") == "s" and v is not None:
                value = shared[int(v.text)]
            elif inline is not None:
                value = "".join(inline.itertext())
            else:
                value = v.text if v is not None and v.text else ""
            if str(value).strip():
                result.append((address, str(value)))
        return result


def export_review(source: Path, output: Path) -> dict[str, int]:
    if output.exists():
        raise FileExistsError("SAFETY STOP: output exists; never overwrite")
    if source.resolve() == output.resolve():
        raise ValueError("SAFETY STOP: source/output collision")
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    rows = []
    counts = {}
    for address, raw in _visible_cells(source):
        record = parse_physiotherapy_cell(raw)
        counts[record.status] = counts.get(record.status, 0) + 1
        # No raw patient-related notes or names from adjacent columns in export.
        rows.append({"source_sheet": "ΠΡΟΓΡΑΜΜΑ", "source_cell": address,
                     **asdict(record)})
    with output.open("x", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(
            stream, fieldnames=["source_sheet", "source_cell",
                                *Interpretation.__dataclass_fields__]
        )
        writer.writeheader()
        writer.writerows(rows)
    if hashlib.sha256(source.read_bytes()).hexdigest() != source_hash:
        raise RuntimeError("SAFETY STOP: source changed")
    return counts


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only column C proposals")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    counts = export_review(args.source, args.output)
    print("MANUAL COLUMN C PROPOSED REVIEW ONLY; NO IMPORT")
    print("Counts:", counts)
    print("Source unchanged: True")


if __name__ == "__main__":
    main()
