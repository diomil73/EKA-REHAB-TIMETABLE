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


def test_master_visual_style_has_matching_a_and_b_clinic_banners():
    assert master_visual_style.FIRST_CLINIC_LABEL == "Α' ΚΛΙΝΙΚΗ"
    assert master_visual_style.SECOND_CLINIC_LABEL == "Β' ΚΛΙΝΙΚΗ"
    source = inspect.getsource(master_visual_style._ensure_clinic_banners)
    assert 'clinic="A"' in source
    assert 'label=FIRST_CLINIC_LABEL' in source
    assert 'clinic="B"' in source
    assert 'label=SECOND_CLINIC_LABEL' in source


def test_clinic_banners_are_larger_and_more_prominent():
    source = inspect.getsource(master_visual_style._style_clinic_banner)
    assert "banner.Font.Size = 14" in source
    assert "banner.Font.Bold = True" in source
    assert "ws.Rows(row).RowHeight = 30" in source


def test_clinic_banner_styling_never_inserts_or_deletes_rows():
    ensure_source = inspect.getsource(master_visual_style._ensure_clinic_banner)
    apply_source = inspect.getsource(master_visual_style._apply_master_visual_style)
    assert ".Insert(" not in ensure_source
    assert ".Delete(" not in ensure_source
    assert ".Insert(" not in apply_source
    assert ".Delete(" not in apply_source
    assert "refusing to move data" in ensure_source


def test_master_body_uses_larger_font_with_cell_level_shrink_to_fit():
    source = inspect.getsource(master_visual_style._apply_master_visual_style)
    assert "body.Font.Size = 10.5" in source
    assert "body.ShrinkToFit = True" in source
    assert "room_cell.Font.Size = 12" in source
    assert "room_cell.Font.Bold = True" in source
    assert "treatment_range.Font.Size = 10.5" in source
    assert "treatment_range.ShrinkToFit = True" in source


def test_patient_identity_typography_preserves_smaller_doctor_line():
    source = inspect.getsource(master_visual_style._format_patient_identity)
    assert "cell.Font.Size = 11.5" in source
    assert "cell.ShrinkToFit = True" in source
    assert "Font.Size = 9.5" in source
    assert "RED_CROSS" in source


def test_unified_builder_refreshes_master_before_navigation_and_visual_style():
    from pathlib import Path

    build_path = Path(__file__).resolve().parents[1] / "Python" / "tools" / "build_unified_preview.py"
    source = build_path.read_text(encoding="utf-8")
    projection_pos = source.index("refresh_master_projection_in_place(output)")
    nav_pos = source.index("finalize_user_navigation(output)")
    style_pos = source.index("apply_master_visual_style_in_place(output)")
    assert projection_pos < nav_pos < style_pos
    assert 'print(f"[OK] MASTER projection refresh ({projected} patients)")' in source
    assert 'print("[OK] MASTER visual baseline")' in source
