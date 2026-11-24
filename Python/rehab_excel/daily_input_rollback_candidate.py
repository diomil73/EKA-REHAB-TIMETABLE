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
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
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
    """Replace journal rows without rewriting Excel's namespace declarations.

    ElementTree's whole-document serialization renames prefixes and can leave
    mc:Ignorable referring to prefixes no longer declared. Excel may reject
    that XML even though an independent Python reader accepts its cells.
    """
    all_rows = (JOURNAL_HEADERS, *rows)
    if len(all_rows) > 1048576:
        raise RollbackCandidateError("Journal worksheet capacity exhausted")
    try:
        root = ET.fromstring(original)
        if root.tag != f"{{{_MAIN_NS}}}worksheet":
            raise RollbackCandidateError("Unexpected journal XML root")
        if root.find(f"{{{_MAIN_NS}}}sheetData") is None:
            raise RollbackCandidateError("Journal sheet data unavailable")

        # Preserve every original byte outside <sheetData> (including all
        # namespace declarations, extLst, mc:Ignorable and XML attributes).
        match = re.search(
            rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)sheetData(?=[\s>])[^<>]*>"
            rb"(?P<body>.*?)</(?P=prefix)sheetData\s*>",
            original, flags=re.DOTALL,
        )
        if match is None or re.search(
            rb"<(?:[A-Za-z_][\w.-]*:)?sheetData(?=[\s>])", original[match.end():]
        ):
            raise RollbackCandidateError("Journal sheetData markup ambiguous")
        prefix = match.group("prefix") or b""
        q = prefix.decode("ascii")

        generated = []
        for row_index, fields in enumerate(all_rows, start=1):
            if len(fields) != len(JOURNAL_HEADERS):
                raise RollbackCandidateError("Journal row width invalid")
            cells = []
            for index, value in enumerate(fields):
                address = f"{chr(65 + index)}{row_index}"
                if row_index > 1 and index == 1:
                    if type(value) is not int or value < 0:
                        raise RollbackCandidateError("Journal sequence invalid")
                    cells.append(
                        f'<{q}c r="{address}" t="n"><{q}v>{value}</{q}v></{q}c>'
                    )
                else:
                    if type(value) is not str:
                        raise RollbackCandidateError("Journal string invalid")
                    # String content only; no untrusted value becomes XML markup.
                    safe = escape(value)
                    cells.append(
                        f'<{q}c r="{address}" t="inlineStr"><{q}is>'
                        f'<{q}t xml:space="preserve">{safe}</{q}t>'
                        f'</{q}is></{q}c>'
                    )
            generated.append(
                f'<{q}row r="{row_index}">{"".join(cells)}</{q}row>'
            )
        result = (
            original[:match.start("body")]
            + "".join(generated).encode("utf-8")
            + original[match.end("body"):]
        )
        # Change only an existing dimension ref, keeping original tag/prefix.
        dimension = re.search(
            rb"<(?:[A-Za-z_][\w.-]*:)?dimension\b[^<>]*>", result
        )
        if dimension is not None:
            item = dimension.group()
            if len(re.findall(rb"\bref=(\x22|\x27)[^\x22\x27]*\1", item)) != 1:
                raise RollbackCandidateError("Journal dimension ref ambiguous")
            patched = re.sub(
                rb"\bref=(\x22|\x27)[^\x22\x27]*\1",
                lambda m: b"ref=" + m.group(1)
                + f"A1:E{len(all_rows)}".encode("ascii") + m.group(1),
                item, count=1,
            )
            result = result[:dimension.start()] + patched + result[dimension.end():]
        verified = ET.fromstring(result)
        sheet = verified.find(f"{{{_MAIN_NS}}}sheetData")
        if sheet is None or len(sheet.findall(f"{{{_MAIN_NS}}}row")) != len(all_rows):
            raise RollbackCandidateError("Journal XML row readback invalid")
        return result
    except (ET.ParseError, UnicodeError, TypeError, ValueError) as exc:
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
