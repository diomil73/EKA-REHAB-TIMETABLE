import inspect

from rehab_excel import master_visual_style
from rehab_excel.native_excel import excel_rgb


def test_master_visual_palette_has_distinct_pt_family_body_colors():
    colors = master_visual_style.BODY_COLORS
    assert colors["fth"] == excel_rgb(217, 234, 247)
    assert colors["robotic"] == excel_rgb(230, 228, 243)
    assert colors["pool"] == excel_rgb(216, 240, 244)
    assert colors["reclined"] == excel_rgb(232, 239, 246)
    assert len({colors["fth"], colors["robotic"], colors["pool"], colors["reclined"]}) == 4


def test_master_visual_palette_uses_strong_infectious_yellow():
    assert master_visual_style.INFECTIOUS_YELLOW == excel_rgb(255, 216, 77)


def test_master_visual_style_hides_infectious_source_column_without_deleting_it():
    source = inspect.getsource(master_visual_style._apply_master_visual_style)
    assert 'ws.Columns("A").Hidden = True' in source
    assert 'Columns("A").Delete' not in source
    assert '_set_fill(ws, f"B{row}:C{row}", INFECTIOUS_YELLOW)' in source


def test_master_visual_style_uses_compact_card_rows_and_wrap():
    source = inspect.getsource(master_visual_style._apply_master_visual_style)
    assert "ws.Rows(row).RowHeight = 64" in source
    assert "body.WrapText = True" in source
    assert 'ws.Range(f"C{header_row + 1}:C{last_row}").HorizontalAlignment = -4131' in source


def test_master_visual_style_adds_visible_patient_row_separator():
    source = inspect.getsource(master_visual_style._apply_master_visual_style)
    assert "bottom = row_range.Borders(9)" in source
    assert "bottom.Weight = 2" in source
    assert "bottom.Color = ROW_SEPARATOR_GRAY" in source


def test_master_visual_style_preserves_header_body_tone_pairing():
    for key in (
        "fth",
        "robotic",
        "pool",
        "reclined",
        "ergo",
        "logo",
        "efa",
        "psych",
        "afternoon",
    ):
        assert master_visual_style.HEADER_COLORS[key] != master_visual_style.BODY_COLORS[key]


def test_master_visual_style_bounds_visible_scroll_area_from_column_b():
    source = inspect.getsource(master_visual_style._apply_master_visual_style)
    assert 'ws.ScrollArea = f"B1:M{last_row + 1}"' in source


def test_unified_builder_applies_master_visual_style_after_navigation():
    from pathlib import Path

    build_path = Path(__file__).resolve().parents[1] / "Python" / "tools" / "build_unified_preview.py"
    source = build_path.read_text(encoding="utf-8")
    nav_pos = source.index("finalize_user_navigation(output)")
    style_pos = source.index("apply_master_visual_style_in_place(output)")
    assert nav_pos < style_pos
    assert 'print("[OK] MASTER visual baseline")' in source
