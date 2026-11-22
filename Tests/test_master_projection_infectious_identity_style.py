"""MASTER identity fill must not inherit infectious formatting from template."""
from rehab_excel.master_projection_writer import Win32ComMasterProjectionBackend
from rehab_excel.master_visual_style import (
    BODY_COLORS, ROW_SEPARATOR_GRAY, SOFT_VERTICAL_GRAY,
)


class Border:
    def __init__(self, color, weight=2):
        self.Color = color
        self.Weight = weight


class Interior:
    def __init__(self, color):
        self.Color = color


class Range:
    def __init__(self, fill, border_colors):
        self.Interior = Interior(fill)
        self._borders = {edge: Border(color) for edge, color in border_colors.items()}

    def Borders(self, edge):
        return self._borders[edge]


def test_noninfectious_identity_resets_yellow_fill_and_only_yellow_borders():
    backend = Win32ComMasterProjectionBackend()
    yellow = backend.palette.infectious_yellow
    normal_existing = 1234567
    row = Range(yellow, {7: yellow, 8: yellow, 9: yellow, 10: normal_existing})
    backend._apply_noninfectious_identity_style(row)

    assert row.Interior.Color == BODY_COLORS["identity"]
    assert row.Borders(7).Color == SOFT_VERTICAL_GRAY
    assert row.Borders(8).Color == SOFT_VERTICAL_GRAY
    assert row.Borders(9).Color == ROW_SEPARATOR_GRAY
    assert row.Borders(9).Weight == 2
    assert row.Borders(10).Color == normal_existing


def test_noninfectious_identity_works_if_template_is_not_yellow():
    backend = Win32ComMasterProjectionBackend()
    row = Range(9999, {7: 11, 8: 12, 9: 13, 10: 14})
    backend._apply_noninfectious_identity_style(row)
    assert row.Interior.Color == BODY_COLORS["identity"]
    assert [row.Borders(edge).Color for edge in (7, 8, 9, 10)] == [11, 12, 13, 14]


def test_infectious_styling_remains_distinct_and_unchanged():
    backend = Win32ComMasterProjectionBackend()
    row = Range(BODY_COLORS["identity"], {7: 1, 8: 1, 9: 1, 10: 1})
    backend._apply_infectious_style(row)
    yellow = backend.palette.infectious_yellow
    assert row.Interior.Color == yellow
    assert all(row.Borders(edge).Color == yellow for edge in (7, 8, 9, 10))
