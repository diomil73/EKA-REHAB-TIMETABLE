"""Conservative, read-only audit of a DAILY_INPUT *whole-workbook* rollback.

This audits a saved pre-mutation PREPARED snapshot versus a separately saved
restored workbook. Every ZIP member except the exact journal worksheet payload
must be byte-identical; this includes formulas, workbook metadata, styles,
shapes, VBA, UserForms, macros, and all other package resources. Native Excel
can legitimately rewrite unrelated package members during Save, so this strict
check can reject a correct semantic rollback. It must never be relaxed to
ignore unexplained differences, and it is not a clinical or Apply authorization.

The snapshot digest must be bound by a separately verified durable transaction
receipt. A caller-chosen digest is not an authenticity proof. Neither this
module nor its tests perform any workbook writes.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from posixpath import dirname, join, normpath
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook

from .authoritative_commit import file_sha256
from .daily_input_journal import (
    JOURNAL_HEADERS,
    JOURNAL_SHEET,
    DailyInputJournalError,
    decode_operation_journal,
)
from .daily_input_operation import DailyInputOperationPhase as Phase


class FullRollbackAuditError(RuntimeError):
    """A read-only whole-workbook rollback proof failed closed."""


@dataclass(frozen=True)
class FullRollbackAudit:
    snapshot_sha256: str
    restored_sha256: str
    checked_nonjournal_parts: int
    original_journal_revisions: int
    restored_journal_revisions: int


_WORKBOOK_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_WORKSHEET_REL = _REL_DOC_NS + "/worksheet"


def _package_members(path: Path, *, require_vba: bool) -> tuple[dict[str, bytes], str]:
    """Load exact uncompressed ZIP members, and locate the unique journal part."""
    with ZipFile(path, "r") as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise FullRollbackAuditError("Duplicate package members")
        if require_vba and "xl/vbaProject.bin" not in names:
            raise FullRollbackAuditError("VBA project missing")
        parts = {name: archive.read(name) for name in names}
    try:
        book = ET.fromstring(parts["xl/workbook.xml"])
        rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    except (KeyError, ET.ParseError) as exc:
        raise FullRollbackAuditError("Workbook package manifest invalid") from exc

    matches = []
    for sheet in book.findall(f".//{{{_WORKBOOK_NS}}}sheet"):
        if sheet.get("name") == JOURNAL_SHEET:
            matches.append(sheet.get(f"{{{_REL_DOC_NS}}}id"))
    if len(matches) != 1 or not matches[0]:
        raise FullRollbackAuditError("Journal worksheet mapping missing or ambiguous")

    targets = []
    for relationship in rels.findall(f"{{{_REL_PKG_NS}}}Relationship"):
        if relationship.get("Id") != matches[0]:
            continue
        if relationship.get("Type") != _WORKSHEET_REL:
            raise FullRollbackAuditError("Journal mapping has invalid relationship type")
        if relationship.get("TargetMode"):
            raise FullRollbackAuditError("External journal relationship rejected")
        target = relationship.get("Target")
        if not target:
            raise FullRollbackAuditError("Journal mapping target unavailable")
        resolved = normpath(
            target.lstrip("/") if target.startswith("/") else
            join(dirname("xl/workbook.xml"), target)
        )
        if resolved.startswith("../") or resolved not in parts:
            raise FullRollbackAuditError("Journal mapping target invalid")
        targets.append(resolved)
    if len(targets) != 1:
        raise FullRollbackAuditError("Journal worksheet relationship ambiguous")
    return parts, targets[0]


def _read_journal(path: Path) -> tuple:
    workbook = None
    try:
        workbook = load_workbook(path, data_only=False, read_only=False, keep_vba=True)
        if JOURNAL_SHEET not in workbook.sheetnames:
            raise FullRollbackAuditError("Journal worksheet unavailable")
        ws = workbook[JOURNAL_SHEET]
        if ws.sheet_state != "veryHidden" or not ws.protection.sheet:
            raise FullRollbackAuditError("Journal protection or visibility changed")
        if tuple(ws.cell(1, col).value for col in range(1, 6)) != JOURNAL_HEADERS:
            raise FullRollbackAuditError("Journal headers invalid")
        rows = []
        for index in range(2, ws.max_row + 1):
            row = tuple(ws.cell(index, col).value for col in range(1, 6))
            if any(value is None for value in row):
                raise FullRollbackAuditError("Blank or partial journal record")
            if type(row[1]) is float and row[1].is_integer():
                row = (row[0], int(row[1]), *row[2:])
            rows.append(row)
        decode_operation_journal(rows)
        return tuple(rows)
    except FullRollbackAuditError:
        raise
    except (DailyInputJournalError, Exception) as exc:
        raise FullRollbackAuditError("Journal native package readback invalid") from exc
    finally:
        if workbook is not None:
            workbook.close()


def _audit_rollback_package(
    prepared_snapshot: str | Path,
    restored_workbook: str | Path,
    *,
    expected_snapshot_sha256: str,
    require_vba: bool,
    required_final_phase: Phase,
) -> FullRollbackAudit:
    """Compare all nonjournal package bytes and legal append-only rollback history.

    The snapshot must be a complete saved XLSM taken after the durable PREPARED
    revision and before all operational mutation. The restored file may differ
    ONLY in its protected, append-only journal worksheet, which must continue
    from that exact PREPARED prefix to a legal FAILED -> ROLLED_BACK terminal
    state. A corresponding independent native Excel read-only reopen and
    trusted snapshot-receipt check are still separate mandatory barriers.
    """
    source = Path(prepared_snapshot).resolve()
    restored = Path(restored_workbook).resolve()
    if (source == restored or source.suffix.lower() != ".xlsm"
            or restored.suffix.lower() != ".xlsm" or not source.is_file()
            or not restored.is_file()):
        raise FullRollbackAuditError("Distinct existing snapshot and restored XLSM required")
    if (type(expected_snapshot_sha256) is not str
            or len(expected_snapshot_sha256) != 64
            or any(c not in "0123456789abcdef" for c in expected_snapshot_sha256)):
        raise FullRollbackAuditError("Snapshot SHA-256 receipt invalid")
    before_sha = file_sha256(source)
    if before_sha != expected_snapshot_sha256:
        raise FullRollbackAuditError("Prepared snapshot changed since receipt")
    after_sha = file_sha256(restored)

    try:
        source_parts, source_journal = _package_members(source, require_vba=require_vba)
        restored_parts, restored_journal = _package_members(restored, require_vba=require_vba)
    except (BadZipFile, OSError, RuntimeError) as exc:
        raise FullRollbackAuditError("Workbook ZIP package readback failed") from exc

    if source_journal != restored_journal:
        raise FullRollbackAuditError("Journal worksheet package relationship changed")
    if source_parts.keys() != restored_parts.keys():
        raise FullRollbackAuditError("Workbook package member inventory changed")
    checked = 0
    for name, before in source_parts.items():
        if name == source_journal:
            continue
        if before != restored_parts[name]:
            # Do not expose custom metadata, sheet names or potential patient data.
            raise FullRollbackAuditError("Nonjournal workbook resource differs")
        checked += 1

    prefix = _read_journal(source)
    tail = _read_journal(restored)
    if not prefix or len(tail) <= len(prefix) + 1 or tail[:len(prefix)] != prefix:
        raise FullRollbackAuditError("Journal rollback prefix or revision chain invalid")
    if decode_operation_journal(prefix)[-1].phase != Phase.PREPARED:
        raise FullRollbackAuditError("Snapshot must end at durable PREPARED")
    appended = decode_operation_journal(tail)[len(prefix):]
    if (
        appended[-1].phase != required_final_phase
        or (required_final_phase == Phase.ROLLED_BACK
            and not appended[-1].rollback_verified)
        or not any(entry.phase == Phase.FAILED for entry in appended)
    ):
        raise FullRollbackAuditError("Required verified rollback phase record missing")
    if appended[-1].operation_id != decode_operation_journal(prefix)[-1].operation_id:
        raise FullRollbackAuditError("Rollback operation identity changed")
    if file_sha256(source) != before_sha or file_sha256(restored) != after_sha:
        raise FullRollbackAuditError("Workbook changed during read-only rollback audit")
    return FullRollbackAudit(
        snapshot_sha256=before_sha,
        restored_sha256=after_sha,
        checked_nonjournal_parts=checked,
        original_journal_revisions=len(prefix),
        restored_journal_revisions=len(tail),
    )

def audit_restored_business_state_before_rollback(
    prepared_snapshot: str | Path,
    restored_workbook: str | Path,
    *,
    expected_snapshot_sha256: str,
    require_vba: bool = True,
) -> FullRollbackAudit:
    """Pre-ROLLED_BACK barrier requiring exact restored state with latest FAILED.

    This is only the first rollback proof stage. It neither marks rollback
    verified nor authorizes ROLLED_BACK. Journal append must happen afterward
    with independent native Excel read-only reopen and terminal proof.
    """
    return _audit_rollback_package(
        prepared_snapshot, restored_workbook,
        expected_snapshot_sha256=expected_snapshot_sha256,
        require_vba=require_vba,
        required_final_phase=Phase.FAILED,
    )


def audit_full_workbook_rollback(
    prepared_snapshot: str | Path,
    restored_workbook: str | Path,
    *,
    expected_snapshot_sha256: str,
    require_vba: bool = True,
) -> FullRollbackAudit:
    """Final read-only exact-package proof after journal reaches ROLLED_BACK.

    A separate trusted receipt and independent native Excel read-only reopen
    remain mandatory before this may contribute to operational approval.
    """
    return _audit_rollback_package(
        prepared_snapshot, restored_workbook,
        expected_snapshot_sha256=expected_snapshot_sha256,
        require_vba=require_vba,
        required_final_phase=Phase.ROLLED_BACK,
    )
