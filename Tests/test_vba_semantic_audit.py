from dataclasses import replace
from hashlib import sha256
from pathlib import Path

import pytest

from rehab_excel.vba_semantic_audit import (
    DOCUMENT_TYPE,
    VbaComponentEvidence,
    VbaSemanticAuditError,
    assert_vba_semantics_preserved,
    audit_vba_components,
)


class FakeCodeModule:
    def __init__(self, content: str):
        self.content = content
        self.CountOfLines = len(content.splitlines())

    def Lines(self, start, count):
        assert start == 1 and count == self.CountOfLines
        return self.content


class FakeComponent:
    def __init__(self, name, kind, source, export_bytes=b"export", resource=None):
        self.Name = name
        self.Type = kind
        self.CodeModule = FakeCodeModule(source)
        self.export_bytes = export_bytes
        self.resource = resource

    def Export(self, filename):
        path = Path(filename)
        path.write_bytes(self.export_bytes)
        if self.resource is not None:
            path.with_suffix(".frx").write_bytes(self.resource)


class FakeComponents:
    def __init__(self, components):
        self.components = components
        self.Count = len(components)

    def Item(self, index):
        return self.components[index - 1]


class FakeProject:
    def __init__(self, components, protection=0):
        self.Protection = protection
        self.VBComponents = FakeComponents(components)


class FakeBook:
    def __init__(self, components, protection=0):
        self.VBProject = FakeProject(components, protection)


def baseline():
    return FakeBook([
        FakeComponent("modDaily", 1, "Public Sub Apply()\nEnd Sub\n"),
        FakeComponent("Sheet1", DOCUMENT_TYPE, "", b"sheet-export"),
        FakeComponent("frmDaily", 3, "Private Sub Test()\nEnd Sub\n",
                      b"form-designer", resource=b"form-picture-bytes"),
    ])


def test_semantic_export_proves_source_and_form_resources():
    evidence = audit_vba_components(baseline())
    assert len(evidence) == 3
    assert {item.name for item in evidence} == {"modDaily", "Sheet1", "frmDaily"}
    assert next(item for item in evidence if item.name == "frmDaily").companion_digest == sha256(b"form-picture-bytes").hexdigest()


def test_benign_xlsm_binary_reserialization_with_unchanged_components_is_allowed():
    before = audit_vba_components(baseline())
    after = audit_vba_components(FakeBook([
        FakeComponent("modDaily", 1, "Public Sub Apply()\r\nEnd Sub\r\n"),
        FakeComponent("Sheet1", DOCUMENT_TYPE, "", b"sheet-export"),
        FakeComponent("frmDaily", 3, "Private Sub Test()\nEnd Sub\n",
                      b"form-designer", resource=b"form-picture-bytes"),
        FakeComponent("SheetJournal", DOCUMENT_TYPE, "", b"journal-empty-document"),
    ]))
    assert assert_vba_semantics_preserved(
        before, after, journal_codename="SheetJournal", expect_new_document=True,
    ) == 3


@pytest.mark.parametrize("case", ["source", "resource", "missing", "unexpected", "journal-code", "journal-type"])
def test_semantic_audit_fails_on_real_vba_changes(case):
    before = audit_vba_components(baseline())
    components = baseline().VBProject.VBComponents.components
    if case == "source":
        components[0].CodeModule = FakeCodeModule("Public Sub Replace()\nEnd Sub")
    elif case == "resource":
        components[2].resource = b"changed-image"
    elif case == "missing":
        components.pop(0)
    elif case == "unexpected":
        components.append(FakeComponent("UnexpectedModule", 1, "", b"rogue"))
    journal = FakeComponent("SheetJournal", DOCUMENT_TYPE, "", b"journal-empty-document")
    if case == "journal-code":
        journal.CodeModule = FakeCodeModule("Public Sub Run()\nEnd Sub")
    elif case == "journal-type":
        journal.Type = 1
    components.append(journal)
    after = audit_vba_components(FakeBook(components))
    with pytest.raises(VbaSemanticAuditError):
        assert_vba_semantics_preserved(
            before, after, journal_codename="SheetJournal", expect_new_document=True,
        )


def test_existing_journal_requires_no_new_document_component():
    snapshot = audit_vba_components(baseline())
    assert assert_vba_semantics_preserved(
        snapshot, snapshot, journal_codename="SheetJournal", expect_new_document=False
    ) == 3
    with pytest.raises(VbaSemanticAuditError, match="addition"):
        assert_vba_semantics_preserved(
            snapshot, snapshot, journal_codename="SheetJournal", expect_new_document=True
        )


def test_parser_or_export_failure_stops_audit():
    book = baseline()
    book.VBProject.VBComponents.components[0].Export = lambda path: None
    with pytest.raises(VbaSemanticAuditError, match="VBA component export missing"):
        audit_vba_components(book)


def test_protected_vba_project_stops_audit():
    with pytest.raises(VbaSemanticAuditError, match="Protected"):
        audit_vba_components(FakeBook([], protection=1))


def test_unexpected_component_type_stops_audit():
    with pytest.raises(VbaSemanticAuditError, match="Unknown VBA"):
        audit_vba_components(FakeBook([FakeComponent("Invalid", 999, "")]))


def test_duplicate_component_name_stops_audit():
    with pytest.raises(VbaSemanticAuditError, match="duplicate"):
        audit_vba_components(FakeBook([
            FakeComponent("Same", 1, ""),
            FakeComponent("same", 1, ""),
        ]))
