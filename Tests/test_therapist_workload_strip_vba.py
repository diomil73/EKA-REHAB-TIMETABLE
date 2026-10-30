import inspect

from rehab_excel import therapist_workload_strip_vba as module


def test_workload_module_remains_stable_backend_dependency():
    assert module.MODULE_NAME == "modTherapistWorkloadStrip"
    assert module.SHAPE_PREFIX == "ekaWorkload_"
    assert module.BUTTON_NAME == "ekaWorkload_Refresh"
    assert "Public Function TherapistCurrentWorkload" in module.MODULE_CODE
    assert "Public Function TherapistWorkloadCapacity" in module.MODULE_CODE
    assert 'ThisWorkbook.Worksheets("THERAPIST_DAILY")' in module.MODULE_CODE


def test_workload_backend_counts_unique_occupied_timeslots():
    code = module.MODULE_CODE

    assert 'Set providerTimes = CreateObject("Scripting.Dictionary")' in code
    assert "providerTimes(timeValue) = True" in code
    assert "TherapistCurrentWorkload = providerTimes.Count" in code
    assert "DEFAULT_CAPACITY As Long = 6" in code
    assert "TherapistWorkloadCapacity = DEFAULT_CAPACITY" in code


def test_workload_backend_merges_matching_therapist_across_layout_blocks():
    code = module.MODULE_CODE

    assert "providerKey = NormalizeProviderName(providerName)" in code
    assert "headerKey = NormalizeProviderName" in code
    assert "If headerKey = providerKey Then" in code
    assert "headerRows = Array(1, 11)" in code
    assert "firstRows = Array(2, 12)" in code
    assert "lastRows = Array(8, 18)" in code
    assert "For col = 2 To 10" in code


def test_workload_module_no_longer_renders_or_buttons_on_master():
    code = module.MODULE_CODE
    install_source = inspect.getsource(module.install_therapist_workload_strip)

    assert 'Worksheets("MASTER_SCHEDULE")' not in code
    assert "AddShape" not in code
    assert "RenderTherapistWorkloadStrip" not in code
    assert "ShowTherapistWorkload" not in code
    assert "AddShape" not in install_source
    assert "_delete_legacy_master_workload_shapes(workbook)" in install_source


def test_installer_removes_legacy_workload_shapes_by_prefix():
    source = inspect.getsource(module._delete_legacy_master_workload_shapes)

    assert 'workbook.Worksheets("MASTER_SCHEDULE")' in source
    assert "name.startswith(SHAPE_PREFIX)" in source
    assert "master.Shapes(name).Delete()" in source
