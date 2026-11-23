"""Build a separate, unpromoted rollback candidate from a PREPARED snapshot.

The entire package, including VBA, UserForms, styles, formulas and all
nonjournal structures, comes from the pre-mutation snapshot. ONLY the journal
worksheet XML is reconstructed from the verified append-only FAILED history.
Use inline strings to avoid stale sharedStrings references across snapshots.

This helper never opens Excel, modifies its inputs, overwrites an existing
output, promotes the candidate, or records ROLLED_BACK. It is NOT an
operational recovery coordinator or clinical authorization. A trusted
snapshot receipt, independent native Excel reopen, journal writer,
final ROLLED_BACK proof, and real transaction integration remain mandatory.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from .authoritative_commit import file_sha256
from .daily_input_full_rollback_audit import (
    FullRollbackAuditError,
    _package_members,
    _read_journal,
    audit_restored_business_state_before_rollback,
)
from .daily_input_journal import JOURNAL_HEADERS, decode_operation_journal
from .daily_input_operation import DailyInputOperationPhase as Phase


class RollbackCandidateError(RuntimeError):
    """The independent rollback candidate could not be safely constructed."""


@dataclass(frozen=True)
class RollbackCandidate:
    candidate_path: Path
    candidate_sha256: str
    source_snapshot_sha256: str
    failed_workbook_sha256: str
    nonjournal_parts_preserved: int


_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _journal_xml_for_history(original: bytes, rows: tuple[tuple, ...]) -> bytes:
    try:
        root = ET.fromstring(original)
        sheet_data = root.find(f"{{{_MAIN_NS}}}sheetData")
        if sheet_data is None:
            raise RollbackCandidateError("Journal sheet data unavailable")
        # Clear ONLY journal row content; preserve worksheet properties,
        # protection, visibility mapping and other non-row metadata.
        for child in list(sheet_data):
            sheet_data.remove(child)

        all_rows = (JOURNAL_HEADERS, *rows)
        if len(all_rows) > 1048576:
            raise RollbackCandidateError("Journal worksheet capacity exhausted")
        for row_index, fields in enumerate(all_rows, start=1):
            row = ET.SubElement(
                sheet_data, f"{{{_MAIN_NS}}}row", {"r": str(row_index)}
            )
            for index, value in enumerate(fields):
                address = f"{chr(65 + index)}{row_index}"
                if row_index > 1 and index == 1:
                    cell = ET.SubElement(
                        row, f"{{{_MAIN_NS}}}c", {"r": address, "t": "n"}
                    )
                    ET.SubElement(cell, f"{{{_MAIN_NS}}}v").text = str(value)
                else:
                    cell = ET.SubElement(
                        row, f"{{{_MAIN_NS}}}c", {"r": address, "t": "inlineStr"}
                    )
                    is_node = ET.SubElement(cell, f"{{{_MAIN_NS}}}is")
                    ET.SubElement(is_node, f"{{{_MAIN_NS}}}t").text = str(value)

        dimensions = root.findall(f"{{{_MAIN_NS}}}dimension")
        for dimension in dimensions:
            dimension.set("ref", f"A1:E{len(all_rows)}")
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    except (ET.ParseError, TypeError, ValueError) as exc:
        raise RollbackCandidateError("Journal XML reconstruction invalid") from exc


def build_rollback_candidate(
    prepared_snapshot: str | Path,
    failed_workbook: str | Path,
    candidate_path: str | Path,
    *,
    expected_snapshot_sha256: str,
    require_vba: bool = True,
) -> RollbackCandidate:
    """Write a NEW separate candidate only, then audit it read-only and fail closed."""
    snapshot = Path(prepared_snapshot).resolve()
    failed = Path(failed_workbook).resolve()
    candidate = Path(candidate_path).resolve()

    if (len({snapshot, failed, candidate}) != 3
            or any(path.suffix.casefold() != ".xlsm" for path in (snapshot, failed, candidate))
            or not snapshot.is_file() or not failed.is_file() or candidate.exists()):
        raise RollbackCandidateError("Distinct source files and unused XLSM output required")
    if (type(expected_snapshot_sha256) is not str
            or len(expected_snapshot_sha256) != 64
            or any(c not in "0123456789abcdef" for c in expected_snapshot_sha256)):
        raise RollbackCandidateError("Expected snapshot receipt invalid")
    snapshot_sha = file_sha256(snapshot)
    if snapshot_sha != expected_snapshot_sha256:
        raise RollbackCandidateError("Snapshot changed since trusted receipt")
    failed_sha = file_sha256(failed)

    created = False
    try:
        source_parts, source_journal = _package_members(snapshot, require_vba=require_vba)
        _, failed_journal = _package_members(failed, require_vba=require_vba)
        if source_journal != failed_journal:
            raise RollbackCandidateError("Journal worksheet location changed")
        source_rows = _read_journal(snapshot)
        failed_rows = _read_journal(failed)
        if not source_rows or failed_rows[:len(source_rows)] != source_rows:
            raise RollbackCandidateError("Failed journal does not retain PREPARED prefix")
        if decode_operation_journal(source_rows)[-1].phase != Phase.PREPARED:
            raise RollbackCandidateError("Snapshot is not at PREPARED boundary")
        if (len(failed_rows) <= len(source_rows)
                or decode_operation_journal(failed_rows)[-1].phase != Phase.FAILED
                or failed_rows[-1][0] != source_rows[-1][0]):
            raise RollbackCandidateError("No durable FAILED operation to recover")
        if file_sha256(snapshot) != snapshot_sha or file_sha256(failed) != failed_sha:
            raise RollbackCandidateError("Input workbook changed during preparation")

        replacement = _journal_xml_for_history(
            source_parts[source_journal], failed_rows,
        )
        # Exclusive creation prevents clobbering any user file. Partial
        # output is removed on synchronous failures, never either input.
        with ZipFile(candidate, "x") as output:
            created = True
            with ZipFile(snapshot, "r") as source_zip:
                for info in source_zip.infolist():
                    payload = replacement if info.filename == source_journal else source_parts[info.filename]
                    output.writestr(info, payload)
        verified = audit_restored_business_state_before_rollback(
            snapshot, candidate,
            expected_snapshot_sha256=expected_snapshot_sha256,
            require_vba=require_vba,
        )
        if file_sha256(failed) != failed_sha or file_sha256(snapshot) != snapshot_sha:
            raise RollbackCandidateError("Input workbook changed during candidate construction")
        return RollbackCandidate(
            candidate_path=candidate,
            candidate_sha256=file_sha256(candidate),
            source_snapshot_sha256=snapshot_sha,
            failed_workbook_sha256=failed_sha,
            nonjournal_parts_preserved=verified.checked_nonjournal_parts,
        )
    except Exception as exc:
        if created:
            try:
                candidate.unlink()
            except OSError:
                pass
        if isinstance(exc, RollbackCandidateError):
            raise
        raise RollbackCandidateError("Rollback candidate creation or independent audit failed") from exc
