from openpyxl import Workbook

from rehab_excel.package_probe import WorkbookPackageProbe


def test_package_probe_preserves_literal_types_and_identifies_formula_errors(tmp_path):
    path = tmp_path / "typed-package.xlsx"
    workbook = Workbook()
    ws = workbook.active
    ws.title = "SOURCE"
    ws["A1"] = "001"
    ws["A2"] = 1
    ws["A3"] = True
    ws["A4"] = "=1+1"
    ws["A5"] = "#N/A"
    ws["A5"].data_type = "e"
    workbook.save(path)
    workbook.close()

    probe = WorkbookPackageProbe(path)
    values = probe.typed_cell_values("SOURCE")
    assert values["A1"] == "001"
    assert values["A2"] == 1
    assert type(values["A2"]) is int
    assert values["A3"] is True
    assert type(values["A3"]) is bool
    assert probe.formula_cells("SOURCE") == {"A4": "1+1"}
    assert probe.error_cells("SOURCE") == {"A5": "#N/A"}
