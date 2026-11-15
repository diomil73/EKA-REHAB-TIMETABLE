import inspect

from rehab_excel.native_excel import Win32ComExcelBackend


def test_rich_text_patch_resets_stale_whole_cell_font_before_semantic_runs():
    source = inspect.getsource(Win32ComExcelBackend._apply_patch)
    assert "if patch.text_runs:" in source
    assert "if patch.font_role is None:" in source
    assert "cell.Font.Color = self.palette.default_black" in source
    assert "run_color = self._font_color_for_role(run.font_role, self.palette)" in source
