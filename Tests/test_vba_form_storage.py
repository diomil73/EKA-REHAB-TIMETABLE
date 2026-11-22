from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from rehab_excel.vba_form_storage import (
    VbaFormStorageError, assert_form_storage_preserved, form_storage_evidence,
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
