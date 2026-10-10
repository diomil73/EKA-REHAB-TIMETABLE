from datetime import date, time

from rehab_core.models import Session
from rehab_excel.master_toolbar_vba import MASTER_TOOLBAR_MODULE_CODE
from rehab_excel.native_excel import NativeStylePalette, excel_rgb
from rehab_excel.render_plan import _is_robotic_session


def test_robotic_treatment_label_is_styled_robotic_even_without_flag():
    session = Session(
        "S1",
        "P1",
        "T1",
        date(2026, 10, 10),
        time(10, 0),
        treatment="Ρομποτικό",
        robotic=False,
    )

    assert _is_robotic_session(session) is True


def test_outpatient_fill_is_readable_muted_green_not_marker_green():
    palette = NativeStylePalette()
    assert palette.outpatient_light_blue == excel_rgb(198, 224, 180)
    assert palette.outpatient_light_blue != excel_rgb(0, 255, 0)


def test_therapist_daily_navigation_reapplies_application_shell():
    start = MASTER_TOOLBAR_MODULE_CODE.index("Public Sub ToolbarTherapistDaily()")
    end = MASTER_TOOLBAR_MODULE_CODE.index("End Sub", start)
    procedure = MASTER_TOOLBAR_MODULE_CODE[start:end]

    assert 'Worksheets("THERAPIST_DAILY").Activate' in procedure
    assert "KeepApplicationShell" in procedure


def test_absence_navigation_also_reapplies_application_shell():
    start = MASTER_TOOLBAR_MODULE_CODE.index("Public Sub ToolbarAbsences()")
    end = MASTER_TOOLBAR_MODULE_CODE.index("End Sub", start)
    procedure = MASTER_TOOLBAR_MODULE_CODE[start:end]

    assert 'Worksheets("DAILY_INPUT").Activate' in procedure
    assert "KeepApplicationShell" in procedure
