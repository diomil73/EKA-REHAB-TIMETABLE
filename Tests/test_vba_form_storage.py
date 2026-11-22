from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from rehab_excel.vba_form_storage import (
    VbaFormStorageError, assert_form_storage_preserved, compare_form_storage_evidence, stream_change_categories, form_storage_evidence,
    assert_matching_source_relative_change_footprints,
)


class FakeOle:
    def __init__(self, streams, issues=None):
        self.streams = streams
        self.parsing_issues = issues or []

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def listdir(self, *, streams, storages):
        assert streams and not storages
        return [path.split("/") for path in self.streams]

    def openstream(self, parts):
        return BytesIO(self.streams["/".join(parts)])


def workbook(path):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"synthetic binary")


def test_form_ole_storage_roundtrip_is_order_independent(tmp_path):
    path = tmp_path / "demo.xlsm"
    workbook(path)
    ole = lambda _: FakeOle({
        "FormOne/f": b"form state",
        "FormOne/o": b"object state",
        "FormOne/i00/f": b"nested form stream",
        "Other/Stream": b"irrelevant",
    })
    snapshot = form_storage_evidence(path, ("FormOne",), ole_factory=ole)
    assert len(snapshot) == 1
    assert len(snapshot[0][1]) == 3
    reverse = lambda _: FakeOle({
        "Other/Stream": b"irrelevant changed",
        "FormOne/i00/f": b"nested form stream",
        "FormOne/o": b"object state",
        "FormOne/f": b"form state",
    })
    later = form_storage_evidence(path, ("FormOne",), ole_factory=reverse)
    assert assert_form_storage_preserved(snapshot, later) == 1


@pytest.mark.parametrize("change", ["changed", "missing", "extra"])
def test_form_storage_refuses_real_stream_changes(tmp_path, change):
    path = tmp_path / "demo.xlsm"
    workbook(path)
    initial = form_storage_evidence(path, ("FormOne",), ole_factory=lambda _: FakeOle({
        "FormOne/f": b"form-state",
        "FormOne/o": b"objects",
    }))
    contents = {"FormOne/f": b"form-state", "FormOne/o": b"objects"}
    if change == "changed":
        contents["FormOne/o"] = b"changed-objects"
    elif change == "missing":
        contents.pop("FormOne/o")
    else:
        contents["FormOne/i00/f"] = b"unexpected"
    later = form_storage_evidence(path, ("FormOne",), ole_factory=lambda _: FakeOle(contents))
    with pytest.raises(VbaFormStorageError, match="mismatch"):
        assert_form_storage_preserved(initial, later)


def test_missing_required_storage_or_form_stream_fails_closed(tmp_path):
    path = tmp_path / "demo.xlsm"
    workbook(path)
    for streams in ({}, {"FormOne/o": b"objects"}):
        with pytest.raises(VbaFormStorageError):
            form_storage_evidence(path, ("FormOne",), ole_factory=lambda _: FakeOle(streams))


def test_parser_issue_fails_closed(tmp_path):
    path = tmp_path / "demo.xlsm"
    workbook(path)
    with pytest.raises(VbaFormStorageError, match="structural issues"):
        form_storage_evidence(path, ("FormOne",),
                              ole_factory=lambda _: FakeOle({"FormOne/f": b"x"}, issues=["corrupt"]))


def test_duplicate_case_insensitive_stream_names_fail_closed(tmp_path):
    path = tmp_path / "demo.xlsm"
    workbook(path)
    with pytest.raises(VbaFormStorageError, match="Ambiguous"):
        form_storage_evidence(path, ("FormOne",), ole_factory=lambda _: FakeOle({
            "FormOne/f": b"1", "FormOne/F": b"2",
        }))


def test_missing_zip_project_fails_closed(tmp_path):
    path = tmp_path / "demo.xlsm"
    with ZipFile(path, "w") as archive:
        archive.writestr("note.txt", b"no VBA")
    with pytest.raises(VbaFormStorageError, match="Unable to parse"):
        form_storage_evidence(path, ("FormOne",), ole_factory=lambda _: FakeOle({}))


def test_embedded_stream_diagnostic_counts_only_without_identifiers():
    baseline = (("privateForm", (("f", "a"), ("o", "b"), ("control/old", "c"))),)
    saved = (
        ("privateForm", (("f", "changed"), ("o", "b"), ("control/new", "d"))),
        ("otherForm", (("f", "z"),)),
    )
    stats = compare_form_storage_evidence(baseline, saved)
    assert stats == {
        "forms_missing": 0, "forms_added": 1, "streams_missing": 1,
        "streams_added": 1, "streams_changed": 1, "forms_with_changed_streams": 1,
    }
    with pytest.raises(VbaFormStorageError) as caught:
        assert_form_storage_preserved(baseline, saved)
    diagnostic = str(caught.value)
    assert "streams_changed=1" in diagnostic
    assert "privateForm" not in diagnostic
    assert "otherForm" not in diagnostic
    assert "control/new" not in diagnostic


def test_embedded_stream_diagnostic_rejects_duplicate_form_keys():
    with pytest.raises(VbaFormStorageError, match="Ambiguous"):
        compare_form_storage_evidence(
            (("form", (("f", "a"),)), ("form", (("f", "b"),))),
            (("form", (("f", "a"),)),),
        )


def test_embedded_stream_diagnostic_rejects_duplicate_stream_keys():
    with pytest.raises(VbaFormStorageError, match="Ambiguous"):
        compare_form_storage_evidence(
            (("form", (("f", "a"), ("f", "b"))),),
            (("form", (("f", "a"),)),),
        )


def test_stream_categories_are_private_and_consistent():
    old = (("private", (("f", "a"), ("o", "b"), ("custom/path", "c"))),)
    new = (("private", (("f", "changed"), ("o", "changed"), ("custom/path", "changed"))),)
    assert stream_change_categories(old, new) == {
        "form_f_changed": 1, "form_o_changed": 1, "other_streams_changed": 1,
    }


def test_source_relative_stream_footprint_detects_same_count_different_stream():
    source = (("privateForm", (("f", "a"), ("o", "b"), ("other", "c"))),)
    noedit = (("privateForm", (("f", "a1"), ("o", "b1"), ("other", "c"))),)
    journal = (("privateForm", (("f", "a2"), ("o", "b"), ("other", "c2"))),)
    assert compare_form_storage_evidence(noedit, journal)["streams_changed"] == 2
    with pytest.raises(VbaFormStorageError, match="footprints differ") as caught:
        assert_matching_source_relative_change_footprints(source, noedit, journal)
    assert "privateForm" not in str(caught.value)


def test_source_relative_stream_footprints_match_without_waiving_raw_byte_mismatch():
    source = (("privateForm", (("f", "a"), ("o", "b"))),)
    noedit = (("privateForm", (("f", "a1"), ("o", "b1"))),)
    journal = (("privateForm", (("f", "a2"), ("o", "b2"))),)
    assert assert_matching_source_relative_change_footprints(source, noedit, journal) == 2
    with pytest.raises(VbaFormStorageError, match="mismatch"):
        assert_form_storage_preserved(noedit, journal)


@pytest.mark.parametrize("bad", [
    (("privateForm", (("f", "a2"),)),),
    (("privateForm", (("f", "a2"), ("f", "a3"), ("o", "b2"))),),
])
def test_source_relative_footprint_missing_or_duplicate_stream_fails_closed(bad):
    source = (("privateForm", (("f", "a"), ("o", "b"))),)
    noedit = (("privateForm", (("f", "a2"), ("o", "b2"))),)
    with pytest.raises(VbaFormStorageError):
        assert_matching_source_relative_change_footprints(source, noedit, bad)
