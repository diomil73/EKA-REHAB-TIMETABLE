from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from rehab_excel.package_probe import WorkbookPackageProbe


CANDIDATE_BASELINES = {
    "Excel/previews/UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm": {
        "size": 731415,
        "sha256": "1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13",
    },
    "Excel/previews/UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm": {
        "size": 622070,
        "sha256": "fa9765da7079137bcb1ff935c3d08bc84b409c8335f9cd719f3583b89e4fc67d",
    },
    "Excel/previews/UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm": {
        "size": 629797,
        "sha256": "4df997b6a5e5b0b56b37127ee00dd8710cc26c7c118f1742d58c9969bcbe7afa",
    },
}

MASTER_HEADERS = {
    "B": "Θάλαμος",
    "C": "Ασθενής",
    "K": "ΨΥΧΟΛΟΓΟΙ",
    "L": "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ",
    "M": "Κατάσταση",
}
PLANNER_HEADERS = ("PatientID", "Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος")


class SchemaComparisonError(RuntimeError):
    pass


@dataclass(frozen=True)
class CandidateSchemaResult:
    relative_path: str
    sha256: str
    size: int
    master_header_row: int | None
    master_matches: tuple[tuple[str, bool], ...]
    planner_positions: tuple[tuple[str, tuple[int, ...]], ...]
    vba_present: bool

    @property
    def master_ok(self) -> bool:
        return self.master_header_row is not None and all(ok for _, ok in self.master_matches)

    @property
    def planner_ok(self) -> bool:
        return all(len(columns) == 1 for _, columns in self.planner_positions)

    @property
    def schema_match(self) -> bool:
        return self.master_ok and self.planner_ok


def file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _normalized(value: object) -> str:
    return "" if value is None else str(value).strip().casefold()


def find_master_header_row(cells: dict[str, object]) -> int | None:
    for row in range(1, 11):
        if (
            _normalized(cells.get(f"B{row}")) == _normalized(MASTER_HEADERS["B"])
            and _normalized(cells.get(f"C{row}")) == _normalized(MASTER_HEADERS["C"])
        ):
            return row
    return None


def has_vba_project_member(names: set[str]) -> bool:
    """Find the expected VBA project ZIP member case-insensitively."""
    return "xl/vbaproject.bin" in {name.casefold() for name in names}


def inspect_candidate(root: Path, relative_path: str) -> CandidateSchemaResult:
    baseline = CANDIDATE_BASELINES[relative_path]
    path = (root / relative_path).resolve()
    if not path.is_file():
        raise SchemaComparisonError(f"EXPECTED_CANDIDATE_MISSING: {relative_path}")

    before_size = path.stat().st_size
    before_hash = file_sha256(path)
    if before_size != baseline["size"] or before_hash != baseline["sha256"]:
        raise SchemaComparisonError(f"FROZEN_BASELINE_MISMATCH: {relative_path}")

    try:
        with ZipFile(path, "r") as archive:
            corrupt_member = archive.testzip()
            if corrupt_member is not None:
                raise SchemaComparisonError("ZIP_INTEGRITY_FAILURE")
            names = set(archive.namelist())
            if not has_vba_project_member(names):
                raise SchemaComparisonError("VBA_PROJECT_MISSING")
    except (BadZipFile, OSError) as exc:
        raise SchemaComparisonError(f"PACKAGE_READ_FAILURE: {type(exc).__name__}") from exc

    probe = WorkbookPackageProbe(path)
    if "MASTER_SCHEDULE" not in probe.sheet_names or "PATIENT_PLANNER" not in probe.sheet_names:
        raise SchemaComparisonError("REQUIRED_SHEET_MISSING")

    master = probe.cell_values("MASTER_SCHEDULE")
    planner = probe.cell_values("PATIENT_PLANNER")
    header_row = find_master_header_row(master)
    master_matches = tuple(
        (
            column,
            header_row is not None
            and _normalized(master.get(f"{column}{header_row}")) == _normalized(expected),
        )
        for column, expected in MASTER_HEADERS.items()
    )

    planner_positions = []
    for expected in PLANNER_HEADERS:
        columns = tuple(
            column_index
            for address, value in planner.items()
            if address.endswith("1")
            and address[:-1].isalpha()
            and _normalized(value) == _normalized(expected)
            for column_index in (address[:-1],)
        )
        planner_positions.append((expected, columns))

    after_size = path.stat().st_size
    after_hash = file_sha256(path)
    if (before_size, before_hash) != (after_size, after_hash):
        raise SchemaComparisonError(f"SOURCE_CHANGED_DURING_READ: {relative_path}")

    return CandidateSchemaResult(
        relative_path=relative_path,
        sha256=before_hash,
        size=before_size,
        master_header_row=header_row,
        master_matches=master_matches,
        planner_positions=tuple(planner_positions),
        vba_present=True,
    )
