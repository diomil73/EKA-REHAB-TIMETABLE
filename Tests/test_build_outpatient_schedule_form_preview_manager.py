from __future__ import annotations

import inspect

from tools import build_outpatient_schedule_form_preview as module


def test_outpatient_form_preview_installs_manager_after_form_and_edit_support():
    source = inspect.getsource(module.main)

    form_pos = source.index("create_outpatient_schedule_form_preview(")
    edit_pos = source.index("install_outpatient_schedule_edit_support_in_workbook(")
    manager_pos = source.index("install_outpatient_schedule_manager_in_workbook(")

    assert form_pos < edit_pos < manager_pos


def test_outpatient_form_preview_exposes_manager_installer():
    assert callable(module.install_outpatient_schedule_manager_in_workbook)
