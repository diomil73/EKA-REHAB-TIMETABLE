import inspect
from pathlib import Path

from openpyxl import Workbook, load_workbook

from rehab_core.models import Patient
from rehab_excel.master_projection import MasterProjectionRow
from rehab_excel.master_projection_writer import (
    MasterProjectionWriteError,
    Win32ComMasterProjectionBackend,
    create_master_projection_preview,
    refresh_master_projection_in_place,
    _build_psychology_display,
)


class _FakeBackend:
    def apply(self, workbook_path, projection):
        wb = load_workbook(workbook_path, keep_vba=False)
        try:
            ws = wb["MASTER_SCHEDULE"]
            header_row = 1
            for candidate in range(1, min(ws.max_row, 10) + 1):
                if (
                    str(ws.cell(candidate, 2).value or "").strip() == "Θάλαμος"
                    and str(ws.cell(candidate, 3).value or "").strip() == "Ασθενής"
                ):
                    header_row = candidate
                    break
            expanded = str(ws.cell(header_row, 13).value or "").strip() == "Κατάσταση"
            status_col = 13 if expanded else 11
            total_cols = 13 if expanded else 11

            for item in projection:
                row = item.target_row
                ws.cell(row, 1).value = f"=PATIENT_PLANNER!D{item.planner_row}"
                ws.cell(row, 2).value = f"=PATIENT_PLANNER!B{item.planner_row}"
                ws.cell(row, 3).value = f"=PATIENT_PLANNER!C{item.planner_row}"
                ws.cell(row, 4).value = f"=PATIENT_PLANNER!F{item.planner_row}"
                ws.cell(row, 5).value = f"=PATIENT_PLANNER!I{item.planner_row}"
                ws.cell(row, 6).value = f"=PATIENT_PLANNER!N{item.planner_row}"
                ws.cell(row, 7).value = f"=PATIENT_PLANNER!L{item.planner_row}"
                ws.cell(row, 8).value = f"=PATIENT_PLANNER!P{item.planner_row}"
                ws.cell(row, 9).value = f"=PATIENT_PLANNER!R{item.planner_row}"
                ws.cell(row, 10).value = f"=PATIENT_PLANNER!T{item.planner_row}"
                ws.cell(row, status_col).value = f"=PATIENT_PLANNER!E{item.planner_row}"
                if expanded and "PATIENT_PLANNER" in wb.sheetnames:
                    planner = wb["PATIENT_PLANNER"]
                    headers = {
                        str(value).strip().casefold(): index
                        for index, value in enumerate(
                            next(planner.iter_rows(min_row=1, max_row=1, values_only=True)),
                            start=1,
                        )
                        if value is not None
                    }
                    def planner_value(header):
                        col = headers.get(header.casefold())
                        return planner.cell(item.planner_row, col).value if col else None
                    ws.cell(row, 11).value = _build_psychology_display(
                        planner_value("Ψυχ_Ώρα"),
                        planner_value("Ψυχ_Ημέρες"),
                        planner_value("Ψυχ_Ψυχολόγος"),
                    )
            tail_start = max((item.target_row for item in projection), default=1) + 1
            for row in range(tail_start, ws.max_row + 1):
                for col in range(1, total_cols + 1):
                    ws.cell(row, col).value = None
            wb.save(workbook_path)
        finally:
            wb.close()


def _book(tmp_path: Path) -> Path:
    path = tmp_path / "source.xlsm"
    wb = Workbook()
    patients = wb.active
    patients.title = "PATIENTS"
    patients.append([
        "PatientID",
        "Θάλαμος",
        "Ασθενής",
        "Μολυσματικός",
        "Κατάσταση",
        "PatientType",
    ])
    patients.append([1, "A09", "ΧΑΣΙΚΟΣ", "Ο", "παρών", "Εσωτερικός"])
    patients.append([2, "A02", "ΒΑΡΒΑΡΑΣ", "Ν", "παρών", "Εσωτερικός"])
    patients.append([3, "A09", "ΤΕΣΤ", "Ν", "παρών", "Εσωτερικός"])

    settings = wb.create_sheet("SETTINGS")
    settings["H1"] = "ROOMS"
    settings["H2"] = "A01"
    settings["H3"] = "A02"
    settings["H4"] = "A09"

    planner = wb.create_sheet("PATIENT_PLANNER")
    planner.append([
        "PatientID",
        "Θάλαμος",
        "Ασθενής",
        "Μολυσματικός",
        "Κατάσταση",
        "ΦΘ_Ώρα",
        "ΦΘ_Ημέρες",
        "ΦΘ_Θεραπευτής",
    ])
    for row in range(2, 5):
        planner.cell(row, 1).value = f"=PATIENTS!A{row}"
        planner.cell(row, 2).value = f"=PATIENTS!B{row}"
        planner.cell(row, 3).value = f"=PATIENTS!C{row}"
        planner.cell(row, 4).value = f"=PATIENTS!D{row}"
        planner.cell(row, 5).value = f"=PATIENTS!E{row}"

    master = wb.create_sheet("MASTER_SCHEDULE")
    master.append([
        "Μολυσματικός",
        "Θάλαμος",
        "Ασθενής",
        "ΦΘ",
        "Ρομποτικό",
        "Πισίνα",
        "Ανακλινόμενο",
        "Εργο",
        "Λογο",
        "ΕΦΑ",
        "Κατάσταση",
    ])
    for row in range(2, 6):
        master.cell(row, 1).value = f"=PATIENT_PLANNER!D{row}"
        master.cell(row, 2).value = f"=PATIENT_PLANNER!B{row}"
        master.cell(row, 3).value = f"=PATIENT_PLANNER!C{row}"
        master.cell(row, 4).value = f"=PATIENT_PLANNER!F{row}"
        master.cell(row, 5).value = f"=PATIENT_PLANNER!I{row}"
        master.cell(row, 6).value = f"=PATIENT_PLANNER!N{row}"
        master.cell(row, 7).value = f"=PATIENT_PLANNER!L{row}"
        master.cell(row, 8).value = f"=PATIENT_PLANNER!P{row}"
        master.cell(row, 9).value = f"=PATIENT_PLANNER!R{row}"
        master.cell(row, 10).value = f"=PATIENT_PLANNER!T{row}"
        master.cell(row, 11).value = f"=PATIENT_PLANNER!E{row}"

    wb.save(path)
    wb.close()
    return path


def test_create_master_projection_preview_preserves_source_and_verifies(tmp_path, monkeypatch):
    source = _book(tmp_path)
    output = tmp_path / "preview.xlsm"

    monkeypatch.setattr(
        "rehab_excel.master_projection_writer.build_master_projection",
        lambda _path: (
            MasterProjectionRow(2, 3, "2", "A02", "ΒΑΡΒΑΡΑΣ", True, "παρών"),
            MasterProjectionRow(3, 4, "3", "A09", "ΤΕΣΤ", True, "παρών"),
            MasterProjectionRow(4, 2, "1", "A09", "ΧΑΣΙΚΟΣ", False, "παρών"),
        ),
    )

    before = source.read_bytes()
    report = create_master_projection_preview(
        source,
        output,
        backend=_FakeBackend(),
        overwrite=True,
    )

    assert report.projected_rows == 3
    assert report.source_unchanged is True
    assert report.verified_in_output is True
    assert source.read_bytes() == before
    assert output.exists()


def test_create_master_projection_preview_rejects_same_source_and_output(tmp_path):
    source = _book(tmp_path)

    try:
        create_master_projection_preview(source, source)
    except MasterProjectionWriteError as exc:
        assert "differ from source" in str(exc)
    else:
        raise AssertionError("Expected same source/output to be rejected")


def test_master_writer_verification_supports_shifted_header(tmp_path, monkeypatch):
    source = _book(tmp_path)
    wb = load_workbook(source)
    try:
        ws = wb["MASTER_SCHEDULE"]
        ws.insert_rows(1, amount=3)
        wb.save(source)
    finally:
        wb.close()

    output = tmp_path / "shifted_preview.xlsm"
    monkeypatch.setattr(
        "rehab_excel.master_projection_writer.build_master_projection",
        lambda _path: (
            MasterProjectionRow(5, 3, "2", "A02", "ΒΑΡΒΑΡΑΣ", True, "παρών"),
            MasterProjectionRow(6, 2, "1", "A09", "ΧΑΣΙΚΟΣ", False, "παρών"),
        ),
    )

    report = create_master_projection_preview(
        source,
        output,
        backend=_FakeBackend(),
        overwrite=True,
    )

    assert report.verified_in_output is True


def test_master_writer_supports_expanded_psychology_afternoon_schema(tmp_path, monkeypatch):
    source = _book(tmp_path)
    wb = load_workbook(source)
    try:
        ws = wb["MASTER_SCHEDULE"]
        ws.insert_cols(11, amount=2)
        ws["K1"] = "ΨΥΧΟΛΟΓΟΙ"
        ws["L1"] = "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ"
        ws["M1"] = "Κατάσταση"
        planner = wb["PATIENT_PLANNER"]
        planner["I1"] = "Ψυχ_Ώρα"
        planner["J1"] = "Ψυχ_Ημέρες"
        planner["K1"] = "Ψυχ_Ψυχολόγος"
        planner["I3"] = 0.5
        planner["J3"] = "ΔΕ, ΤΕ, ΠΑ"
        planner["K3"] = "ΣΥΝΘΕΤΙΚΟΣ ΨΥΧΟΛΟΓΟΣ"
        wb.save(source)
    finally:
        wb.close()

    output = tmp_path / "expanded_preview.xlsm"
    monkeypatch.setattr(
        "rehab_excel.master_projection_writer.build_master_projection",
        lambda _path: (
            MasterProjectionRow(
                2, 3, "2", "A02", "ΒΑΡΒΑΡΑΣ", True, "παρών",
                psychology_display="12:00\nΔΕ, ΤΕ, ΠΑ\nΣΥΝΘΕΤΙΚΟΣ ΨΥΧΟΛΟΓΟΣ",
            ),
        ),
    )

    report = create_master_projection_preview(
        source,
        output,
        backend=_FakeBackend(),
        overwrite=True,
    )

    assert report.verified_in_output is True
    wb = load_workbook(output, data_only=False)
    try:
        ws = wb["MASTER_SCHEDULE"]
        assert ws["K1"].value == "ΨΥΧΟΛΟΓΟΙ"
        assert ws["L1"].value == "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ"
        assert ws["K2"].value == "12:00\nΔΕ, ΤΕ, ΠΑ\nΣΥΝΘΕΤΙΚΟΣ ΨΥΧΟΛΟΓΟΣ"
        assert ws["L2"].value in (None, "")
        assert ws["M2"].value == "=PATIENT_PLANNER!E3"
    finally:
        wb.close()


def test_master_refresh_preserves_existing_visual_shell_formats():
    source = inspect.getsource(Win32ComMasterProjectionBackend.apply)
    assert "if row > old_last:" in source
    assert 'ws.Range(f"A{template_row}:{end_col}{template_row}").Copy()' in source
    assert 'ws.Range(f"B{template_row}:C{template_row}").Copy()' in source
    assert 'ws.Range(f"B{row}:C{row}").PasteSpecial(Paste=-4122)' in source


def test_master_refresh_does_not_force_full_workbook_recalculation():
    source = inspect.getsource(Win32ComMasterProjectionBackend.apply)
    assert "CalculateFull" not in source


def test_date_effective_master_writer_excludes_discharge_after_last_day(tmp_path):
    from datetime import date
    from rehab_excel.patient_exit_source import HEADERS as EXIT_HEADERS
    from rehab_excel.patient_registry_source import read_patient_registry

    source = _book(tmp_path)
    wb = load_workbook(source)
    sheet = wb.create_sheet("PATIENT_EXITS")
    sheet.append(EXIT_HEADERS)
    sheet.append(["2", "inpatient_discharge", date(2026, 11, 16), "SYNTHETIC"])
    wb.save(source)
    wb.close()

    count = refresh_master_projection_in_place(
        source, backend=_FakeBackend(), target_date=date(2026, 11, 17),
    )
    assert count == 2
    wb = load_workbook(source, data_only=False)
    try:
        formulas = [
            wb["MASTER_SCHEDULE"].cell(row, 3).value
            for row in range(2, 7)
        ]
        assert "=PATIENT_PLANNER!C3" not in formulas
        assert "=PATIENT_PLANNER!C4" in formulas
        assert "=PATIENT_PLANNER!C2" in formulas
    finally:
        wb.close()
    assert len(read_patient_registry(source)) == 3

def test_psychology_display_matches_existing_component_order_and_omits_empty_parts():
    assert _build_psychology_display(0.5, "ΔΕ, ΠΕ", "ΔΟΚΙΜΑΣΤΗΣ") == (
        "12:00\nΔΕ, ΠΕ\nΔΟΚΙΜΑΣΤΗΣ"
    )
    assert _build_psychology_display(None, "ΔΕ, ΠΕ", "") == "ΔΕ, ΠΕ"
    assert _build_psychology_display("  ", None, "ΔΟΚΙΜΑΣΤΗΣ") == "ΔΟΚΙΜΑΣΤΗΣ"
    assert _build_psychology_display("μετά τις 16:00", "", "") == "μετά τις 16:00"

def test_expanded_master_last_row_includes_stale_psychology_and_afternoon_cells(monkeypatch):
    class _Cell:
        def __init__(self, row):
            self.Row = row

        def End(self, _direction):
            return self

    class _Rows:
        Count = 100

    class _Worksheet:
        Rows = _Rows()

        def Cells(self, _row, col):
            return _Cell({1: 5, 11: 20, 12: 30, 13: 4}.get(col, 1))

    monkeypatch.setattr(
        Win32ComMasterProjectionBackend,
        "_schema",
        classmethod(lambda _cls, _ws: (13, 13)),
    )
    assert Win32ComMasterProjectionBackend._last_value_row(_Worksheet()) == 30
