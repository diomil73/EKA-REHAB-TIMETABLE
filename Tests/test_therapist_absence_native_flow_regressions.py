import inspect

from rehab_excel.daily_input_action_vba import DAILY_INPUT_MODULE_CODE
from rehab_excel.daily_input_sheet import (
    THERAPIST_ABSENCE_REASONS,
    _configure_lists,
    _format_daily_sheet,
)


def test_daily_input_refresh_does_not_require_retired_master_workload_shape():
    assert 'master.Shapes("ekaWorkload_Refresh")' not in DAILY_INPUT_MODULE_CODE
    assert "RenderTherapistWorkloadStrip" in DAILY_INPUT_MODULE_CODE
    assert "ApplyMasterView" in DAILY_INPUT_MODULE_CODE


def test_daily_input_continues_into_replacement_resolution_after_refresh():
    refresh_position = DAILY_INPUT_MODULE_CODE.index('currentPhase = "refresh operational views"')
    replacements_position = DAILY_INPUT_MODULE_CODE.index(
        'currentPhase = "resolve replacements"'
    )
    assert refresh_position < replacements_position
    assert "ResolveDailyReplacementsInPlace()" in DAILY_INPUT_MODULE_CODE
    assert 'JsonLongValue(responseText, "replacement_sessions_needed")' in DAILY_INPUT_MODULE_CODE


def test_therapist_absence_reason_has_named_dropdown():
    list_source = inspect.getsource(_configure_lists)
    sheet_source = inspect.getsource(_format_daily_sheet)

    assert THERAPIST_ABSENCE_REASONS
    assert "PY_THERAPIST_REASONS" in list_source
    assert "spec.therapist_reasons" in list_source
    assert '=PY_THERAPIST_REASONS' in sheet_source
    assert 'Range(f"E{tr1}:E{tr2}")' in sheet_source
