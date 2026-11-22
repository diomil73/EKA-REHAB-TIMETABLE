"""Read original VBA UserForm OLE storage streams directly from XLSM.

UserForm .frx export bytes can vary even for an unchanged, read-only workbook.
This module compares embedded UserForm streams in xl/vbaProject.bin instead.
Fail closed if form storage is not found, ambiguous, malformed or unreadable.
Never logs names, stream paths, actual bytes or digests.
"""
from __future__ import annotations

from hashlib import sha256
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile


class VbaFormStorageError(RuntimeError):
    pass


def form_storage_evidence(
    workbook_path: Path,
    form_names: tuple[str, ...],
    *,
    ole_factory=None,
) -> tuple[tuple[str, tuple[tuple[str, str], ...]], ...]:
    """Return deterministic raw OLE stream digests for specified UserForms.

    Complete stream sets are compared for each UserForm, not just known f/o
    streams. Storage path is required to be a top-level child by the form name,
    and must contain the required f stream. No best-effort fallback.
    """
    if not form_names or any(not name for name in form_names):
        raise VbaFormStorageError("No unambiguous UserForm inventory supplied")
    if len({name.casefold() for name in form_names}) != len(form_names):
        raise VbaFormStorageError("Duplicate UserForm name")

    if ole_factory is None:
        try:
            import olefile
        except ImportError as exc:
            raise VbaFormStorageError("olefile is required for UserForm storage verification") from exc
        ole_factory = olefile.OleFileIO

    try:
        with ZipFile(workbook_path) as archive:
            vba_bytes = archive.read("xl/vbaProject.bin")
        with ole_factory(BytesIO(vba_bytes)) as ole:
            if hasattr(ole, "parsing_issues") and ole.parsing_issues:
                raise VbaFormStorageError("VBA OLE parser reported structural issues")
            stream_paths = tuple(tuple(str(piece) for piece in path)
                                 for path in ole.listdir(streams=True, storages=False))
            evidence = []
            for name in form_names:
                matches = [(path, path[1:]) for path in stream_paths
                           if len(path) >= 2 and path[0].casefold() == name.casefold()]
                if not matches:
                    raise VbaFormStorageError("Required UserForm storage is missing")
                relative_paths = [tuple(piece.casefold() for piece in relative)
                                  for _, relative in matches]
                if len(relative_paths) != len(set(relative_paths)):
                    raise VbaFormStorageError("Ambiguous UserForm OLE streams")
                if not any(relative == ("f",) for relative in relative_paths):
                    raise VbaFormStorageError("Required UserForm form stream missing")
                streams = []
                for path, relative in matches:
                    data = ole.openstream(list(path)).read()
                    streams.append(("/".join(part.casefold() for part in relative),
                                    sha256(data).hexdigest()))
                evidence.append((name.casefold(), tuple(sorted(streams))))
        return tuple(sorted(evidence))
    except VbaFormStorageError:
        raise
    except Exception as exc:
        raise VbaFormStorageError("Unable to parse VBA UserForm OLE storage") from exc


def compare_form_storage_evidence(before, after) -> dict[str, int]:
    """Count stream-level differences without exposing component IDs or data.

    This is diagnostic only. No mismatch is considered safe or ignored.
    """
    original = dict(before)
    current = dict(after)
    if len(original) != len(before) or len(current) != len(after):
        raise VbaFormStorageError("Ambiguous UserForm evidence identities")
    stats = {
        "forms_missing": len(original.keys() - current.keys()),
        "forms_added": len(current.keys() - original.keys()),
        "streams_missing": 0,
        "streams_added": 0,
        "streams_changed": 0,
        "forms_with_changed_streams": 0,
    }
    for form in original.keys() & current.keys():
        old = dict(original[form])
        new = dict(current[form])
        if len(old) != len(original[form]) or len(new) != len(current[form]):
            raise VbaFormStorageError("Ambiguous UserForm evidence stream identities")
        missing = len(old.keys() - new.keys())
        added = len(new.keys() - old.keys())
        changed = sum(old[key] != new[key] for key in old.keys() & new.keys())
        stats["streams_missing"] += missing
        stats["streams_added"] += added
        stats["streams_changed"] += changed
        if missing or added or changed:
            stats["forms_with_changed_streams"] += 1
    return stats


def assert_form_storage_preserved(before, after) -> int:
    """Require exactly the same UserForm storages and stream bytes."""
    if not before:
        raise VbaFormStorageError("Empty UserForm OLE evidence")
    changes = compare_form_storage_evidence(before, after)
    if any(changes.values()):
        summary = " ".join(f"{key}={value}" for key, value in changes.items())
        raise VbaFormStorageError("Embedded UserForm storage stream mismatch: " + summary)
    return len(before)

def stream_change_categories(before, after) -> dict[str, int]:
    """Diagnostic-only changed stream totals by known OLE stream class.

    The names of custom streams, controls and UserForms are not exposed.
    Each changed stream is grouped into f, o or other.
    """
    original = dict(before)
    current = dict(after)
    if len(original) != len(before) or len(current) != len(after):
        raise VbaFormStorageError("Ambiguous UserForm evidence identities")
    result = {"form_f_changed": 0, "form_o_changed": 0, "other_streams_changed": 0}
    for name in original.keys() & current.keys():
        old = dict(original[name])
        new = dict(current[name])
        if len(old) != len(original[name]) or len(new) != len(current[name]):
            raise VbaFormStorageError("Ambiguous UserForm evidence stream identities")
        for path in old.keys() & new.keys():
            if old[path] != new[path]:
                label = (
                    "form_f_changed" if path == "f" else
                    "form_o_changed" if path == "o" else
                    "other_streams_changed"
                )
                result[label] += 1
    return result


def assert_matching_source_relative_change_footprints(source, noedit, journal) -> int:
    """Compare *which* OLE streams changed relative to the same original.

    Counts alone can mask one stream changing in the no-edit control and a
    different stream changing in the journal trial. This check requires the
    exact source-relative changed-stream identities to match per form.
    It does NOT compare changed stream bytes, prove binary preservation,
    or authorize ignoring strict OLE/resource integrity checks.
    """
    def normalized(snapshot):
        forms = {}
        for name, streams in snapshot:
            if name in forms:
                raise VbaFormStorageError("Ambiguous UserForm evidence identities")
            entries = {}
            for path, digest in streams:
                if path in entries:
                    raise VbaFormStorageError("Ambiguous UserForm evidence stream identities")
                entries[path] = digest
            forms[name] = entries
        return forms

    original = normalized(source)
    left = normalized(noedit)
    right = normalized(journal)
    if not original or original.keys() != left.keys() or original.keys() != right.keys():
        raise VbaFormStorageError("UserForm footprint inventory differs")
    matched_changes = 0
    for name, streams in original.items():
        noedit_streams = left[name]
        journal_streams = right[name]
        if streams.keys() != noedit_streams.keys() or streams.keys() != journal_streams.keys():
            raise VbaFormStorageError("UserForm footprint stream inventory differs")
        changed_noedit = {key for key in streams if streams[key] != noedit_streams[key]}
        changed_journal = {key for key in streams if streams[key] != journal_streams[key]}
        if changed_noedit != changed_journal:
            raise VbaFormStorageError("Journal and no-edit change footprints differ")
        matched_changes += len(changed_noedit)
    return matched_changes
