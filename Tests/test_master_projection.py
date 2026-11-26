import inspect
from pathlib import Path

from openpyxl import Workbook

from rehab_core.models import Patient, PatientType
from rehab_excel.master_projection import (
    _projection_metadata,
    build_master_projection,
    remap_master_formula,
)


def _workbook(tmp_path: Path) -> Path:
    xlsx = tmp_path / "book.xlsx"
    path = tmp_path / "book.xlsm"
    wb = Workbook()
    ws = wb.active
    ws.title = "PATIENTS"
    ws.append(["PatientID", "Θάλαμος", "Ασθενής", "Μολυσματικός", "Κατάσταση"])
    ws.append([1, "A09", "ΧΑΣΙΚΟΣ", "Ο", "παρών"])
    ws.append([2, "A02", "ΒΑΡΒΑΡΑΣ", "Ν", "παρών"])
    ws.append([3, "A09", "ΤΕΣΤ", "Ν", "παρών"])
    ws.append([4, "", "OUT", "", ""])
    wb.create_sheet("SETTINGS")
    wb.save(xlsx)
    path.write_bytes(xlsx.read_bytes())
    return path


def test_projection_metadata_uses_lightweight_ooxml_package_probe():
    source = inspect.getsource(_projection_metadata)

    assert "WorkbookPackageProbe(path)" in source
    assert 'probe.cell_values("PATIENTS")' in source
    assert 'probe.cell_values("SETTINGS")' in source
    assert 'probe.cell_values("MASTER_SCHEDULE")' in source
    assert "load_workbook(" not in source


def test_projection_sorts_rooms_naturally_then_patient_name(tmp_path):
    path = _workbook(tmp_path)
    patients = [
        Patient("1", "ΧΑΣΙΚΟΣ", room="A09"),
        Patient("2", "ΒΑΡΒΑΡΑΣ", room="A02", infectious=True),
        Patient("3", "ΤΕΣΤ", room="A09", infectious=True),
        Patient(
            "4",
            "OUT",
            patient_type=PatientType.OUTPATIENT,
        ),
    ]

    rows = build_master_projection(
        path,
        patients=patients,
        room_order=("A09", "A02", "B01", "A01"),
    )

    assert [(row.target_row, row.planner_row, row.patient_id) for row in rows] == [
        (3, 3, "2"),
        (4, 4, "3"),
        (5, 2, "1"),
    ]
    assert all(row.patient_id != "4" for row in rows)


def test_projection_never_interleaves_b_clinic_inside_a_clinic(tmp_path):
    path = _workbook(tmp_path)
    patients = [
        Patient("1", "A11 PATIENT", room="A11"),
        Patient("2", "B08 PATIENT", room="B08"),
        Patient("3", "A05 PATIENT", room="A05"),
    ]

    rows = build_master_projection(
        path,
        patients=patients,
        room_order=("A11", "B08", "A05"),
    )

    assert [(row.room, row.target_row) for row in rows] == [
        ("A05", 3),
        ("A11", 4),
        ("B08", 6),
    ]


def test_projection_keeps_patient_registry_rows_as_formula_targets(tmp_path):
    path = _workbook(tmp_path)
    patients = [
        Patient("3", "ΤΕΣΤ", room="A09", infectious=True),
        Patient("1", "ΧΑΣΙΚΟΣ", room="A09"),
    ]

    rows = build_master_projection(
        path,
        patients=patients,
        room_order=("A09",),
    )

    assert rows[0].display_name == "ΤΕΣΤ"
    assert rows[0].target_row == 3
    assert rows[0].planner_row == 4
    assert rows[1].display_name == "ΧΑΣΙΚΟΣ"
    assert rows[1].target_row == 4
    assert rows[1].planner_row == 2


def test_projection_reserves_exactly_one_banner_row_at_clinic_transition(tmp_path):
    path = _workbook(tmp_path)
    patients = [
        Patient("1", "ΧΑΣΙΚΟΣ", room="A09"),
        Patient("2", "ΒΑΡΒΑΡΑΣ", room="A02"),
        Patient("3", "ΤΕΣΤ", room="B01"),
    ]

    rows = build_master_projection(
        path,
        patients=patients,
        room_order=("A02", "A09", "B01"),
    )

    assert [(row.room, row.target_row) for row in rows] == [
        ("A02", 3),
        ("A09", 4),
        ("B01", 6),
    ]


def test_remap_master_formula_redirects_all_planner_references():
    formula = (
        '=IF(PATIENT_PLANNER!F26="","",'
        'PATIENT_PLANNER!F26&" "&PATIENT_PLANNER!G26)'
    )

    remapped = remap_master_formula(formula, 100)

    assert "PATIENT_PLANNER!F100" in remapped
    assert "PATIENT_PLANNER!G100" in remapped
    assert "PATIENT_PLANNER!F26" not in remapped


def test_remap_master_formula_handles_absolute_row_reference():
    assert (
        remap_master_formula("=PATIENT_PLANNER!$D$2", 100)
        == "=PATIENT_PLANNER!$D$100"
    )


def test_projection_starts_below_shifted_master_header(tmp_path):
    path = _workbook(tmp_path)

    from openpyxl import load_workbook

    wb = load_workbook(path)
    try:
        ws = wb.create_sheet("MASTER_SCHEDULE")
        ws.append([
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
        ws.insert_rows(1, amount=3)
        wb.save(path)
    finally:
        wb.close()

    patients = [
        Patient("1", "ΧΑΣΙΚΟΣ", room="A09"),
        Patient("2", "ΒΑΡΒΑΡΑΣ", room="A02", infectious=True),
    ]

    rows = build_master_projection(
        path,
        patients=patients,
        room_order=("A02", "A09"),
    )

    assert [row.target_row for row in rows] == [6, 7]


def test_date_effective_master_keeps_last_active_day_then_excludes_discharge(tmp_path):
    from datetime import date
    from openpyxl import load_workbook
    from rehab_excel.patient_exit_source import HEADERS as EXIT_HEADERS

    path = _workbook(tmp_path)
    wb = load_workbook(path)
    exit_sheet = wb.create_sheet("PATIENT_EXITS")
    exit_sheet.append(EXIT_HEADERS)
    exit_sheet.append(["2", "inpatient_discharge", date(2026, 11, 16), "DEMO"])
    wb.save(path)
    wb.close()

    sample = [
        Patient("1", "ΧΑΣΙΚΟΣ", room="A09"),
        Patient("2", "ΒΑΡΒΑΡΑΣ", room="A02", infectious=True),
        Patient("3", "ΤΕΣΤ", room="A09"),
    ]
    original = path.read_bytes()
    undated = build_master_projection(
        path, patients=sample, room_order=("A02", "A09"),
    )
    on_last_day = build_master_projection(
        path, patients=sample, room_order=("A02", "A09"),
        target_date=date(2026, 11, 16),
    )
    day_after = build_master_projection(
        path, patients=sample, room_order=("A02", "A09"),
        target_date=date(2026, 11, 17),
    )
    assert {row.patient_id for row in undated} == {"1", "2", "3"}
    assert on_last_day == undated
    assert {row.patient_id for row in day_after} == {"1", "3"}
    assert {row.planner_row for row in day_after} == {2, 4}
    assert path.read_bytes() == original


def test_date_effective_master_rejects_invalid_service_date(tmp_path):
    from rehab_excel.master_projection import MasterProjectionError

    with __import__("pytest").raises(MasterProjectionError, match="valid target date"):
        build_master_projection(_workbook(tmp_path), target_date="2026-11-17")

def test_expanded_master_projection_derives_psychology_display_from_planner_source(tmp_path):
    from datetime import time
    from openpyxl import load_workbook

    path = _workbook(tmp_path)
    wb = load_workbook(path)
    try:
        planner = wb.create_sheet("PATIENT_PLANNER")
        planner.append(["PatientID", "", "", "", "", "", "", "", "Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος"])
        planner.append([1])
        planner.append([2, None, None, None, None, None, None, None, time(16, 30), "ΔΕ, ΠΕ", "ΣΥΝΘΕΤΙΚΟΣ ΨΥΧΟΛΟΓΟΣ"])
        planner.append([3])
        master = wb.create_sheet("MASTER_SCHEDULE")
        master.append(["", "Θάλαμος", "Ασθενής", "", "", "", "", "", "", "", "ΨΥΧΟΛΟΓΟΙ", "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ", "Κατάσταση"])
        wb.save(path)
    finally:
        wb.close()

    rows = build_master_projection(
        path,
        patients=[Patient("2", "ΒΑΡΒΑΡΑΣ", room="A02")],
    )

    assert len(rows) == 1
    assert rows[0].psychology_display == "16:30\nΔΕ, ΠΕ\nΣΥΝΘΕΤΙΚΟΣ ΨΥΧΟΛΟΓΟΣ"

def test_expanded_master_projection_rejects_ambiguous_or_drifted_planner_source(tmp_path):
    from datetime import time
    import pytest
    from openpyxl import load_workbook
    from rehab_excel.master_projection import MasterProjectionError

    for corruption in ("missing-header", "duplicate-header", "patient-id-drift"):
        path = _workbook(tmp_path)
        wb = load_workbook(path)
        try:
            planner = wb.create_sheet("PATIENT_PLANNER")
            planner.append([
                "PatientID", "", "", "", "", "", "", "",
                "Ψυχ_Ώρα", "Ψυχ_Ημέρες", "Ψυχ_Ψυχολόγος",
            ])
            planner.append([1])
            planner.append([
                2, None, None, None, None, None, None, None,
                time(16, 30), "ΔΕ, ΠΕ", "SYNTHETIC PROVIDER",
            ])
            planner.append([3])
            master = wb.create_sheet("MASTER_SCHEDULE")
            master.append([
                "", "Θάλαμος", "Ασθενής", "", "", "", "", "", "", "",
                "ΨΥΧΟΛΟΓΟΙ", "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ", "Κατάσταση",
            ])
            if corruption == "missing-header":
                planner["J1"] = None
            elif corruption == "duplicate-header":
                planner["L1"] = "Ψυχ_Ώρα"
            else:
                planner["A3"] = 99
            wb.save(path)
        finally:
            wb.close()

        with pytest.raises(MasterProjectionError):
            build_master_projection(
                path,
                patients=[Patient("2", "ΒΑΡΒΑΡΑΣ", room="A02")],
            )
