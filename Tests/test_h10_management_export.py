from __future__ import annotations

from datetime import date
from types import SimpleNamespace
from pathlib import Path
import zipfile

import pytest

from tools.export_management_therapist_summary import (
    department_periods,
    inspect_share_package,
    MODE_LABELS,
)


def fact(day, therapist, planned, actual):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        therapist_id=therapist,
        planned_sessions=planned,
        final_active_sessions=actual,
    )


def test_sum_is_based_on_counts_not_mean_of_therapist_percentages():
    result = department_periods([
        fact("2026-10-29", "Alice Secret", 10, 8),
        fact("2026-10-29", "Bob Secret", 2, 2),
        fact("2026-10-30", "Alice Secret", 4, 5),
    ])
    assert result["daily"] == [
        ("2026-10-29", 12, 10, 10 / 12),
        ("2026-10-30", 4, 5, 1.25),
    ]
    assert result["weekly"] == [("2026-10-26", 16, 15, 15 / 16)]
    assert result["monthly"] == [("2026-10", 16, 15, 15 / 16)]
    assert result["overall"] == [("ΟΛΟ ΤΟ ΙΣΤΟΡΙΚΟ", 16, 15, 15 / 16)]
    assert "Alice Secret" not in str(result)
    assert "Bob Secret" not in str(result)
    assert list(MODE_LABELS) == ["daily", "weekly", "monthly", "overall"]


def test_zero_denominator_is_unavailable_not_fake_zero():
    assert department_periods([fact("2026-10-30", "T", 0, 0)])["overall"] == [
        ("ΟΛΟ ΤΟ ΙΣΤΟΡΙΚΟ", 0, 0, None)
    ]


def test_duplicate_or_negative_current_final_fact_fails_closed():
    row = fact("2026-10-30", "T", 2, 2)
    with pytest.raises(ValueError, match="Duplicate"):
        department_periods([row, row])
    with pytest.raises(ValueError, match="Negative"):
        department_periods([fact("2026-10-30", "T", -1, 0)])


def test_export_package_rejects_any_operational_schema_or_vba(tmp_path):
    filename = tmp_path / "public.xlsx"
    with zipfile.ZipFile(filename, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
    inspect_share_package(filename)
    with zipfile.ZipFile(filename, "a") as z:
        z.writestr("xl/vbaProject.bin", "unsafe")
    with pytest.raises(ValueError, match="VBA"):
        inspect_share_package(filename)


def test_export_package_rejects_hidden_operational_schema(tmp_path):
    filename = tmp_path / "public.xlsx"
    with zipfile.ZipFile(filename, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/workbook.xml", "<workbook>THERAPIST_DAY_HISTORY</workbook>")
    with pytest.raises(ValueError, match="leaked"):
        inspect_share_package(filename)


def test_export_package_rejects_hidden_sheets_and_external_relationships(tmp_path):
    hidden = tmp_path / "hidden.xlsx"
    with zipfile.ZipFile(hidden, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/workbook.xml", '<workbook><sheet state="veryHidden"/></workbook>')
    with pytest.raises(ValueError, match="Hidden"):
        inspect_share_package(hidden)

    linked = tmp_path / "external.xlsx"
    with zipfile.ZipFile(linked, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships><Relationship TargetMode="External" /></Relationships>')
    with pytest.raises(ValueError, match="External"):
        inspect_share_package(linked)


def test_export_package_rejects_embedded_objects_and_extra_sheets(tmp_path):
    embedded = tmp_path / "embedded.xlsx"
    with zipfile.ZipFile(embedded, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/embeddings/oleObject1.bin", "secret")
    with pytest.raises(ValueError, match="embedded"):
        inspect_share_package(embedded)

    extra = tmp_path / "extra.xlsx"
    with zipfile.ZipFile(extra, "w") as z:
        for i in range(1, 6):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
    with pytest.raises(ValueError, match="exactly four"):
        inspect_share_package(extra)


def test_share_export_requires_independent_numeric_reopen_and_source_hash():
    import inspect
    from tools import export_management_therapist_summary as mod
    code = inspect.getsource(mod.export_shareable)
    check = inspect.getsource(mod.verify_management_native)
    assert "source_hash = hashlib.sha256(source.read_bytes()).hexdigest()" in code
    assert "verify_management_native(output, periods)" in code
    assert "inspect_share_package(output)" in code
    assert 'int(book.Worksheets.Count) != 4' in check
    assert "int(ws.ChartObjects().Count)" in check
    assert "abs_tol=1e-12" in check
    assert "book.Close(SaveChanges=False)" in check
    assert "hashlib.sha256(file.read_bytes()).hexdigest() != before_hash" in check


def test_shareable_executive_layout_uses_native_buttons_and_hides_excel_chrome():
    import inspect
    from tools import export_management_therapist_summary as mod
    render = inspect.getsource(mod._write_mode_sheet)
    export = inspect.getsource(mod.export_shareable)
    verifier = inspect.getsource(mod.verify_management_native)
    assert "Shapes.AddShape" in render
    assert "ekaExecutiveMode_" in render
    assert 'SubAddress=f"\'{name}\'!A1"' in render
    assert 'ws.Application.ActiveWindow.DisplayHeadings = False' in render
    assert 'ws.Application.ActiveWindow.DisplayGridlines = False' in render
    assert "excel.ActiveWindow.DisplayWorkbookTabs = False" in export
    assert 'ws.Range("A1:J3").Merge()' in render
    assert "ΑΡΧΙΚΑ ΠΡΟΓΡΑΜΜΑΤΙΣΜΕΝΕΣ" in render
    assert "ΤΕΛΙΚΑ ΠΡΑΓΜΑΤΟΠΟΙΗΜΕΝΕΣ" in render
    assert "ΠΟΣΟΣΤΟ ΥΛΟΠΟΙΗΣΗΣ" in render
    assert 'ws.Range("F15").Left' in render
    assert "row = 16 + offset" in verifier
    assert "total_row = 16 + len(expected_rows)" in verifier
    assert "navigation buttons missing" in verifier
    assert "navigation target mismatch" in verifier
    assert "DisplayWorkbookTabs" in verifier


def test_executive_headline_cards_are_verified_against_aggregates():
    import inspect
    from tools import export_management_therapist_summary as mod
    checker = inspect.getsource(mod.verify_management_native)
    assert 'int(ws.Range("A9").Value2) != planned_total' in checker
    assert 'int(ws.Range("C9").Value2) != delivered_total' in checker
    assert 'headline_ratio = ws.Range("F9").Value2' in checker
    assert "KPI percentage mismatch" in checker
    render = inspect.getsource(mod._write_mode_sheet)
    assert "OnAction" not in render
    assert 'Address="",\n                          SubAddress=' in render

def test_management_export_only_publishes_after_native_and_package_verification(
    tmp_path, monkeypatch
):
    import sys
    from types import SimpleNamespace
    from tools import export_management_therapist_summary as module

    source = tmp_path / "fictional-source.xlsm"
    source.write_bytes(b"FICTIONAL_SOURCE_ORIGINAL")
    target = tmp_path / "management-verified.xlsx"
    events = []
    periods = {
        "daily": [("2026-11-24", 1, 1, 1.0)],
        "weekly": [("2026-11-23", 1, 1, 1.0)],
        "monthly": [("2026-11", 1, 1, 1.0)],
        "overall": [("ΟΛΟ ΤΟ ΙΣΤΟΡΙΚΟ", 1, 1, 1.0)],
    }
    monkeypatch.setattr(module, "report_therapist_history", lambda path: ())
    monkeypatch.setattr(module, "department_periods", lambda rows: periods)
    monkeypatch.setattr(module, "sys", SimpleNamespace(platform="win32"))
    monkeypatch.setattr(module, "_write_mode_sheet", lambda ws, mode, rows: None)

    class FakeSheet:
        def Activate(self):
            pass

    class FakeWorksheets:
        Count = 4
        def __call__(self, key):
            return FakeSheet()

    class FakeBook:
        Worksheets = FakeWorksheets()
        def SaveAs(self, filename, FileFormat):
            assert FileFormat == 51
            assert str(filename).endswith(".xlsx")
            assert not target.exists()
            events.append("save-private")
            from pathlib import Path
            Path(filename).write_bytes(b"FICTIONAL_MACROFREE_STAGED")
        def Close(self, SaveChanges=False):
            pass

    class FakeExcel:
        def __init__(self):
            self.Workbooks = SimpleNamespace(Add=FakeBook)
            self.ActiveWindow = SimpleNamespace()
        def Quit(self):
            pass

    fake_client = SimpleNamespace(DispatchEx=lambda name: FakeExcel())
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    def inspect(staged):
        assert staged.is_file()
        assert staged != target and not target.exists()
        events.append("package")
    def native(staged, expected):
        assert expected is periods and not target.exists()
        events.append("native")
    monkeypatch.setattr(module, "inspect_share_package", inspect)
    monkeypatch.setattr(module, "verify_management_native", native)

    before = source.read_bytes()
    module.export_shareable(source, target)
    assert target.read_bytes() == b"FICTIONAL_MACROFREE_STAGED"
    assert events == ["save-private", "package", "native", "package"]
    assert source.read_bytes() == before
    assert not tuple(tmp_path.glob(".diomil_unverified_*"))


def test_management_export_native_reopen_failure_never_leaves_shareable_file(
    tmp_path, monkeypatch
):
    from types import SimpleNamespace
    import sys
    import pytest
    from tools import export_management_therapist_summary as module

    source = tmp_path / "fictional-source.xlsm"
    source.write_bytes(b"FICTIONAL_SOURCE")
    target = tmp_path / "must-not-publish.xlsx"
    data = {mode: [("FICTIONAL", 1, 1, 1.0)] for mode in module.MODES}
    monkeypatch.setattr(module, "report_therapist_history", lambda path: ())
    monkeypatch.setattr(module, "department_periods", lambda rows: data)
    monkeypatch.setattr(module, "sys", SimpleNamespace(platform="win32"))
    monkeypatch.setattr(module, "_write_mode_sheet", lambda ws, mode, rows: None)

    class FakeBook:
        def __init__(self):
            self.Worksheets = SimpleNamespace(Count=4, __call__=lambda *a: None)
        def SaveAs(self, filename, FileFormat):
            from pathlib import Path
            Path(filename).write_bytes(b"UNVERIFIED PRIVATE EXPORT")
        def Close(self, SaveChanges=False):
            pass
    book = FakeBook()
    class FakeSheets:
        Count = 4
        def __call__(self, index):
            return SimpleNamespace(Activate=lambda: None)
    book.Worksheets = FakeSheets()
    excel = SimpleNamespace(
        Workbooks=SimpleNamespace(Add=lambda: book),
        ActiveWindow=SimpleNamespace(),
        Quit=lambda: None,
    )
    client = SimpleNamespace(DispatchEx=lambda arg: excel)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(module, "inspect_share_package", lambda path: None)
    def fail_native(path, expected):
        raise RuntimeError("FICTIONAL PRIVATE FAILED READBACK")
    monkeypatch.setattr(module, "verify_management_native", fail_native)

    with pytest.raises(RuntimeError, match="FAILED READBACK"):
        module.export_shareable(source, target)
    assert not target.exists()
    assert not tuple(tmp_path.glob(".diomil_unverified_*"))
    assert source.read_bytes() == b"FICTIONAL_SOURCE"
