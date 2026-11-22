"""Fail-closed semantic VBA component audit for native Excel journal provisioning.

A new worksheet necessarily adds a document VBComponent, so xl/vbaProject.bin
is not expected to remain byte-identical. The auditor exports every preexisting
component before/after, compares code and designer resources, and allows exactly
one empty worksheet document module for the journal. It never modifies VB code.

The caller must disable macros before Workbooks.Open, hold the expected project
manifest only in memory, and compare after closing and read-only reopening.
"""
from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory

DOCUMENT_TYPE = 100
FORM_TYPE = 3


class VbaSemanticAuditError(RuntimeError):
    """VBA component source/resources cannot be proven unchanged."""


@dataclass(frozen=True)
class VbaComponentEvidence:
    name: str
    component_type: int
    code_digest: str
    export_digest: str
    companion_digest: str | None


def _code_text(component) -> str:
    module = component.CodeModule
    count = int(module.CountOfLines)
    if count < 0:
        raise VbaSemanticAuditError("Invalid VBA module line count")
    if count == 0:
        return ""
    return str(module.Lines(1, count)).replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")


def _digest(path: Path) -> str:
    return sha256(path.read_bytes()).hexdigest()


def audit_vba_components(workbook) -> tuple[VbaComponentEvidence, ...]:
    """Capture source and exported resources for every current VBA component.

    Fail closed if the project is protected/inaccessible or a component export
    fails. Exported form .frx resources are compared separately.
    """
    try:
        project = workbook.VBProject
        if int(project.Protection) != 0:
            raise VbaSemanticAuditError("Protected VBA project cannot be audited")
        components = project.VBComponents
        count = int(components.Count)
        if count < 1:
            raise VbaSemanticAuditError("Empty VBA project")
        entries = []
        with TemporaryDirectory(prefix="diomil_vba_audit_") as directory:
            root = Path(directory)
            for index in range(1, count + 1):
                component = components.Item(index)
                name = str(component.Name)
                kind = int(component.Type)
                if not name or any(entry.name.casefold() == name.casefold() for entry in entries):
                    raise VbaSemanticAuditError("Missing or duplicate VBA component identity")
                source_text = _code_text(component)
                suffix = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}.get(kind)
                if suffix is None:
                    raise VbaSemanticAuditError("Unknown VBA component type")
                export = root / f"component_{index}{suffix}"
                component.Export(str(export))
                if not export.is_file():
                    raise VbaSemanticAuditError("VBA component export missing")
                resource = export.with_suffix(".frx")
                if kind != FORM_TYPE and resource.exists():
                    raise VbaSemanticAuditError("Unexpected VBA component resource")
                entries.append(VbaComponentEvidence(
                    name=name,
                    component_type=kind,
                    code_digest=sha256(source_text.encode("utf-8")).hexdigest(),
                    export_digest=_digest(export),
                    companion_digest=_digest(resource) if resource.exists() else None,
                ))
        return tuple(sorted(entries, key=lambda entry: entry.name.casefold()))
    except VbaSemanticAuditError:
        raise
    except Exception as exc:
        raise VbaSemanticAuditError("VBA project inspection/export failed") from exc


def assert_vba_semantics_preserved(
    before: tuple[VbaComponentEvidence, ...],
    after: tuple[VbaComponentEvidence, ...],
    *,
    journal_codename: str | None,
    expect_new_document: bool,
) -> int:
    """Require original components untouched plus one empty journal document.

    Returns the original component count on success. Any additional component,
    missing original, changed source or resource fails closed.
    """
    previous = {entry.name.casefold(): entry for entry in before}
    current = {entry.name.casefold(): entry for entry in after}
    if (not previous or len(previous) != len(before)
            or len(current) != len(after)):
        raise VbaSemanticAuditError("Missing or duplicate VBA component identities")
    if any(current.get(key) != item for key, item in previous.items()):
        raise VbaSemanticAuditError("Preexisting VBA component code/resource changed")
    added = set(current) - set(previous)
    if bool(added) != expect_new_document:
        raise VbaSemanticAuditError("Unexpected journal component addition/missing addition")
    if added:
        if len(added) != 1 or not journal_codename or added != {journal_codename.casefold()}:
            raise VbaSemanticAuditError("Unexpected added VBA component")
        created = current[next(iter(added))]
        if created.component_type != DOCUMENT_TYPE:
            raise VbaSemanticAuditError("Journal VBA component is not a worksheet document")
        # Empty module is checked by comparing the known empty-code SHA.
        if created.code_digest != sha256(b"").hexdigest():
            raise VbaSemanticAuditError("Journal document contains unexpected VBA code")
    return len(previous)
