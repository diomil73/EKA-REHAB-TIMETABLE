from rehab_excel.package_probe import WorkbookPackageProbe


def test_join_part_normalizes_package_absolute_target_for_zip_member():
    assert (
        WorkbookPackageProbe._join_part(
            "xl/workbook.xml",
            "/xl/worksheets/sheet1.xml",
        )
        == "xl/worksheets/sheet1.xml"
    )


def test_join_part_preserves_relative_relationship_resolution():
    assert (
        WorkbookPackageProbe._join_part(
            "xl/workbook.xml",
            "worksheets/sheet1.xml",
        )
        == "xl/worksheets/sheet1.xml"
    )


def test_join_part_handles_parent_relative_relationship_target():
    assert (
        WorkbookPackageProbe._join_part(
            "xl/worksheets/sheet1.xml",
            "../drawings/drawing1.xml",
        )
        == "xl/drawings/drawing1.xml"
    )
