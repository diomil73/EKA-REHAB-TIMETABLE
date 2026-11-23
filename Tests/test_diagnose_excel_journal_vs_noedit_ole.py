from pathlib import Path
from zipfile import ZipFile

import pytest

from Python.tools import diagnose_excel_journal_vs_noedit_ole as probe
from rehab_excel.vba_semantic_audit import VbaComponentEvidence


class FakeBook:
    ReadOnly = False
    VBASigned = False

    def __init__(self):
        self.saves = 0
        self.closed = []
        self.Worksheets = FakeSheets()

    def Save(self):
        self.saves += 1

    def Close(self, **kwargs):
        self.closed.append(kwargs)


class FakeSheets:
    Count = 1
    def __call__(self, index):
        return type("Sheet", (), {"Name": "Sheet1"})()


class FakeApp:
    def __init__(self):
        self.book = FakeBook()
        self.Workbooks = self
        self.open_args = []
        self.quits = 0

    def Open(self, path, **kwargs):
        self.open_args.append(kwargs)
        # A native read-only reopen must be distinguishable from writable Save.
        self.book.ReadOnly = bool(kwargs.get("ReadOnly", False))
        return self.book

    def Quit(self):
        self.quits += 1


def make_xlsm(path):
    with ZipFile(path, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"fake-vba")


def evidence():
    return (VbaComponentEvidence("frmSynthetic", 3, "code", "export", "frx"),)


def test_two_separate_copies_produce_private_comparison_and_clean_up(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    original = source.read_bytes()
    apps = []
    def factory():
        app = FakeApp()
        apps.append(app)
        return app
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    calls = []
    def install(book):
        calls.append(book)
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", install)
    stream_sets = iter([
        (("frmsynthetic", (("f", "same"), ("o", "same"))),),  # source
        (("frmsynthetic", (("f", "same"), ("o", "same"))),),  # no-edit
        (("frmsynthetic", (("f", "same"), ("o", "same"))),),  # journal
    ])
    monkeypatch.setattr(probe, "form_storage_evidence", lambda path, names: next(stream_sets))
    result = probe.compare_noedit_and_journal_saves(source, excel_factory=factory)
    assert len(apps) == 2
    assert apps[0].book.saves == 1
    assert apps[1].book.saves == 0
    assert len(calls) == 1
    assert all(a.book.closed == [{"SaveChanges": False}] for a in apps)
    assert all(a.quits == 1 for a in apps)
    assert all(a.open_args[0]["ReadOnly"] is False for a in apps)
    assert result["differences"]["streams_changed"] == 0
    assert source.read_bytes() == original
    assert list(tmp_path.iterdir()) == [source]


def test_comparison_detects_new_journal_stream_changes_without_approval(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", lambda wb: None)
    stream_sets = iter([
        (("frmsynthetic", (("f", "same"), ("o", "original"))),),  # source
        (("frmsynthetic", (("f", "same"), ("o", "control"))),),   # no-edit
        (("frmsynthetic", (("f", "same"), ("o", "changed"))),),   # journal
    ])
    monkeypatch.setattr(probe, "form_storage_evidence", lambda path, names: next(stream_sets))
    result = probe.compare_noedit_and_journal_saves(source, excel_factory=FakeApp)
    assert result["differences"]["streams_changed"] == 1
    assert result["categories"]["form_o_changed"] == 1


def test_comparison_refuses_signed_project_and_cleans_up(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    apps = []
    def factory():
        app = FakeApp()
        app.book.VBASigned = True
        apps.append(app)
        return app
    with pytest.raises(probe.SaveComparisonError, match="signed VBA"):
        probe.compare_noedit_and_journal_saves(source, excel_factory=factory)
    assert len(apps) == 1
    assert apps[0].book.closed == [{"SaveChanges": False}]
    assert apps[0].quits == 1
    assert list(tmp_path.iterdir()) == [source]


def test_control_only_saves_two_workbooks_without_installing_journal(tmp_path, monkeypatch):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    original = source.read_bytes()
    apps = []
    def factory():
        app = FakeApp()
        apps.append(app)
        return app
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    def forbidden_install(book):
        raise AssertionError("Control-only must never install a journal")
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", forbidden_install)
    streams = iter([
        (("frmsynthetic", (("f", "original"), ("o", "original"))),),  # source
        (("frmsynthetic", (("f", "before"), ("o", "before"))),),  # first no-edit
        (("frmsynthetic", (("f", "after"), ("o", "after"))),),
    ])
    monkeypatch.setattr(probe, "form_storage_evidence", lambda path, forms: next(streams))
    result = probe.compare_noedit_and_journal_saves(
        source, excel_factory=factory, control_only=True,
    )
    assert result["comparison"] == "NO-EDIT vs NO-EDIT"
    assert result["differences"]["streams_changed"] == 2
    assert result["categories"] == {
        "form_f_changed": 1, "form_o_changed": 1, "other_streams_changed": 0,
    }
    assert len(apps) == 2
    assert [app.book.saves for app in apps] == [1, 1]
    assert all(app.book.closed == [{"SaveChanges": False}] for app in apps)
    assert all(app.quits == 1 for app in apps)
    assert source.read_bytes() == original
    assert list(tmp_path.iterdir()) == [source]


def _designer_signatures(*digests):
    from rehab_excel.vba_form_designer_audit import UserFormDesignerEvidence
    return iter([
        (UserFormDesignerEvidence("frmsynthetic", digest, 4),)
        for digest in digests
    ])


def test_native_designer_readback_matches_both_isolated_reopens(tmp_path, monkeypatch):
    """Each copy reopens read-only; no form is instantiated or source saved."""
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    before = source.read_bytes()
    apps = []

    def factory():
        instance = FakeApp()
        apps.append(instance)
        return instance

    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(
        probe, "audit_userform_designers",
        lambda wb, *, expected_form_names: next(signatures),
    )
    monkeypatch.setattr(
        probe, "form_storage_evidence",
        lambda path, names: (("frmsynthetic", (("f", "A"), ("o", "B"))),),
    )
    def forbidden_journal(workbook):
        raise AssertionError("Control-only must not create a journal")
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", forbidden_journal)
    signatures = _designer_signatures("same", "same", "same", "same")

    result = probe.compare_noedit_and_journal_saves(
        source, excel_factory=factory, control_only=True,
        designer_readback=True,
    )
    assert result["designer_forms_verified"] == 1
    assert result["differences"]["streams_changed"] == 0
    assert len(apps) == 2
    assert [len(app.open_args) for app in apps] == [2, 2]
    assert all([args["ReadOnly"] for args in app.open_args] == [False, True] for app in apps)
    assert all(app.book.closed == [{"SaveChanges": False}, {"SaveChanges": False}] for app in apps)
    assert [app.book.saves for app in apps] == [1, 1]
    assert all(app.quits == 1 for app in apps)
    assert source.read_bytes() == before
    assert list(tmp_path.iterdir()) == [source]


@pytest.mark.parametrize("signatures,expected_stage", [
    (("baseline", "mutated", "baseline", "baseline"), "save no-edit control"),
    (("baseline", "baseline", "other", "other"), "cross-save UserForm designer"),
])
def test_native_designer_delta_refuses_trial_without_mutating_source(
    tmp_path, monkeypatch, signatures, expected_stage,
):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    original = source.read_bytes()
    iterator = _designer_signatures(*signatures)
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(
        probe, "audit_userform_designers",
        lambda wb, *, expected_form_names: next(iterator),
    )
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", lambda wb: None)
    monkeypatch.setattr(
        probe, "form_storage_evidence",
        lambda path, names: (("frmsynthetic", (("f", "same"),)),),
    )
    with pytest.raises(probe.SaveComparisonError, match=expected_stage) as caught:
        probe.compare_noedit_and_journal_saves(
            source, excel_factory=FakeApp, designer_readback=True,
        )
    assert "frmsynthetic" not in str(caught.value)
    assert source.read_bytes() == original
    assert list(tmp_path.iterdir()) == [source]


def test_source_relative_footprint_mismatch_stops_even_with_same_total_change_count(
    tmp_path, monkeypatch,
):
    """Distinct changed paths must not pass because aggregate counts match."""
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    baseline_bytes = source.read_bytes()
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", lambda wb: None)
    # Source f/o unchanged, while each saved copy changes one different path.
    stream_sets = iter([
        (("frmsynthetic", (("f", "original-f"), ("o", "original-o"))),),
        (("frmsynthetic", (("f", "noedit-f"), ("o", "original-o"))),),
        (("frmsynthetic", (("f", "original-f"), ("o", "journal-o"))),),
    ])
    monkeypatch.setattr(probe, "form_storage_evidence", lambda path, names: next(stream_sets))
    with pytest.raises(probe.SaveComparisonError, match="source-relative OLE change footprints") as caught:
        probe.compare_noedit_and_journal_saves(source, excel_factory=FakeApp)
    assert "frmsynthetic" not in str(caught.value)
    assert source.read_bytes() == baseline_bytes
    assert list(tmp_path.iterdir()) == [source]


@pytest.mark.parametrize("sequential,count", [(False, 3), (True, 6)])
def test_native_fictional_journal_trial_stages_reopens_and_never_touches_source(
    tmp_path, monkeypatch, sequential, count,
):
    """Opt-in synthetic revision path does not run VBA or save original."""
    from types import SimpleNamespace
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    original = source.read_bytes()
    apps = []
    staged = []

    class TrialSheets(FakeSheets):
        def __init__(self, book):
            self.book = book

        def __call__(self, index):
            if index == probe.JOURNAL_SHEET:
                return self.book.journal
            return super().__call__(index)

    def factory():
        app = FakeApp()
        app.book.Worksheets = TrialSheets(app.book)
        apps.append(app)
        return app

    def install(book):
        book.journal = SimpleNamespace(
            Name=probe.JOURNAL_SHEET, Visible=2, ProtectContents=True,
        )
        return book.journal

    def stage(book, operation):
        assert book.ReadOnly is False
        assert book.journal.ProtectContents is True
        staged.append(operation)
        return True

    rows = tuple(("FICTIONAL", i, "FAKE", "digest", "payload") for i in range(count))
    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", install)
    monkeypatch.setattr(probe, "append_native_operation_revision", stage)
    monkeypatch.setattr(probe, "read_native_journal_rows", lambda ws: rows)
    monkeypatch.setattr(
        probe, "audit_userform_designers",
        lambda wb, *, expected_form_names: next(signatures),
    )
    signatures = _designer_signatures("same", "same", "same", "same")
    monkeypatch.setattr(
        probe, "form_storage_evidence",
        lambda path, forms: (("frmsynthetic", (("f", "same"), ("o", "same"))),),
    )

    result = probe.compare_noedit_and_journal_saves(
        source, excel_factory=factory, designer_readback=True,
        trial_fictional_revisions=True,
        trial_sequential_operations=sequential,
    )
    assert result["fictional_revisions_verified"] == count
    expected = ["PREPARED", "FAILED", "ROLLED_BACK"]
    assert [operation.phase.value for operation in staged] == (
        expected * (2 if sequential else 1)
    )
    if sequential:
        assert staged[0].operation_id != staged[3].operation_id
        assert staged[0].source_sha256 != staged[3].source_sha256
        assert staged[3].revision == 0
    assert len(apps) == 2
    assert [app.book.saves for app in apps] == [1, 1]
    assert all([args["ReadOnly"] for args in app.open_args] == [False, True] for app in apps)
    assert source.read_bytes() == original
    assert list(tmp_path.iterdir()) == [source]


def test_fictional_journal_trial_requires_independent_journal_designer_readback(
    tmp_path,
):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    with pytest.raises(probe.SaveComparisonError, match="requires journal and designer"):
        probe.compare_noedit_and_journal_saves(
            source, excel_factory=FakeApp,
            control_only=True, designer_readback=True,
            trial_fictional_revisions=True,
        )
    with pytest.raises(probe.SaveComparisonError, match="requires journal and designer"):
        probe.compare_noedit_and_journal_saves(
            source, excel_factory=FakeApp,
            trial_fictional_revisions=True,
        )
    assert list(tmp_path.iterdir()) == [source]


def test_sequential_trial_requires_explicit_fictional_revision_mode(tmp_path):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    before = source.read_bytes()
    with pytest.raises(probe.SaveComparisonError, match="requires fictional revisions"):
        probe.compare_noedit_and_journal_saves(
            source, excel_factory=FakeApp, designer_readback=True,
            trial_sequential_operations=True,
        )
    assert source.read_bytes() == before
    assert list(tmp_path.iterdir()) == [source]


def test_native_prepared_handoff_is_appended_after_two_terminal_chains_and_reopened(
    tmp_path, monkeypatch,
):
    from types import SimpleNamespace
    from pathlib import Path

    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    original = source.read_bytes()
    apps = []
    staged = []
    verified = []

    class TrialSheets(FakeSheets):
        def __init__(self, book):
            self.book = book

        def __call__(self, index):
            if index == probe.JOURNAL_SHEET:
                return self.book.journal
            return super().__call__(index)

    def factory():
        app = FakeApp()
        app.book.Worksheets = TrialSheets(app.book)
        apps.append(app)
        return app

    def install(book):
        book.journal = SimpleNamespace(
            Name=probe.JOURNAL_SHEET, Visible=2, ProtectContents=True,
        )
        return book.journal

    def stage(book, operation):
        assert book.ReadOnly is False
        assert book.journal.ProtectContents is True
        staged.append(operation)
        return True

    def rows(_ws):
        return tuple(
            ("FICTIONAL", index, "FAKE", "digest", "payload")
            for index in range(len(staged))
        )

    handoff = {
        "status": "PREPARED_NOT_PERSISTED",
        "operation_id": "FICTIONAL-PREPARED-7",
        "service_date": "2000-01-03",
        "source_sha256": "a" * 64,
        "payload_sha256": "b" * 64,
        "existing_journal_rows": 6,
        "journal_prefix_sha256": "c" * 64,
        "prepared_row": [
            "FICTIONAL-PREPARED-7", 0, "PREPARED", "d" * 64, "fictional-payload",
        ],
    }
    payload_path = Path("fictional-payload.tsv")

    def prepare(copy):
        assert copy.exists()
        assert len(staged) == 6
        assert apps[1].book.saves == 1
        return {"handoff": handoff, "payload_path": payload_path}

    def verify(copy, metadata, payload):
        assert copy.name == "journal.xlsm"
        assert metadata == handoff
        assert payload == payload_path
        # The writable Excel session must be closed before independent readback.
        assert apps[1].book.closed == [{"SaveChanges": False}]
        verified.append(True)

    monkeypatch.setattr(probe, "audit_vba_components", lambda wb: evidence())
    monkeypatch.setattr(probe, "ensure_daily_input_journal_sheet", install)
    monkeypatch.setattr(probe, "append_native_operation_revision", stage)
    monkeypatch.setattr(probe, "read_native_journal_rows", rows)
    monkeypatch.setattr(probe, "_create_fictional_prepared_handoff", prepare)
    monkeypatch.setattr(probe, "verify_prepared_reopen", verify)
    monkeypatch.setattr(
        probe, "audit_userform_designers",
        lambda wb, *, expected_form_names: next(signatures),
    )
    signatures = _designer_signatures("same", "same", "same", "same")
    monkeypatch.setattr(
        probe, "form_storage_evidence",
        lambda path, forms: (("frmsynthetic", (("f", "same"), ("o", "same"))),),
    )

    result = probe.compare_noedit_and_journal_saves(
        source, excel_factory=factory, designer_readback=True,
        trial_fictional_revisions=True,
        trial_sequential_operations=True,
        trial_prepared_handoff=True,
    )
    assert result["fictional_revisions_verified"] == 6
    assert result["prepared_handoff_verified"] is True
    assert result["prepared_journal_rows"] == 7
    assert len(staged) == 7
    assert [item.phase.value for item in staged[:6]] == [
        "PREPARED", "FAILED", "ROLLED_BACK",
        "PREPARED", "FAILED", "ROLLED_BACK",
    ]
    assert staged[6].operation_id == handoff["operation_id"]
    assert staged[6].phase.value == "PREPARED" and staged[6].revision == 0
    assert len(verified) == 1
    assert len(apps) == 2
    assert [app.book.saves for app in apps] == [1, 2]
    assert all(app.AutomationSecurity == 3 for app in apps)
    assert all(app.EnableEvents is False for app in apps)
    assert all(
        [args["ReadOnly"] for args in app.open_args] == [False, True]
        for app in apps
    )
    assert all(
        app.book.closed == [{"SaveChanges": False}, {"SaveChanges": False}]
        for app in apps
    )
    assert all(app.quits == 1 for app in apps)
    assert source.read_bytes() == original
    assert list(tmp_path.iterdir()) == [source]


def test_prepared_handoff_trial_cannot_skip_prior_terminal_and_designer_gates(tmp_path):
    source = tmp_path / "demo.xlsm"
    make_xlsm(source)
    before = source.read_bytes()
    with pytest.raises(probe.SaveComparisonError, match="requires sequential fictional"):
        probe.compare_noedit_and_journal_saves(
            source, excel_factory=FakeApp, designer_readback=True,
            trial_prepared_handoff=True,
        )
    with pytest.raises(probe.SaveComparisonError, match="requires sequential fictional"):
        probe.compare_noedit_and_journal_saves(
            source, excel_factory=FakeApp, trial_fictional_revisions=True,
            trial_sequential_operations=True, trial_prepared_handoff=True,
        )
    assert source.read_bytes() == before
    assert list(tmp_path.iterdir()) == [source]
