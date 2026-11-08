from __future__ import annotations

from datetime import date
from types import SimpleNamespace
from pathlib import Path
import zipfile

import pytest

from tools.export_management_therapist_summary import (
    department_periods,
    inspect_share_package,
    MODE_LABELS,
)


def fact(day, therapist, planned, actual):
    return SimpleNamespace(
        service_date=date.fromisoformat(day),
        therapist_id=therapist,
        planned_sessions=planned,
        final_active_sessions=actual,
    )


def test_sum_is_based_on_counts_not_mean_of_therapist_percentages():
    result = department_periods([
        fact("2026-10-29", "Alice Secret", 10, 8),
        fact("2026-10-29", "Bob Secret", 2, 2),
        fact("2026-10-30", "Alice Secret", 4, 5),
    ])
    assert result["daily"] == [
        ("2026-10-29", 12, 10, 10 / 12),
        ("2026-10-30", 4, 5, 1.25),
    ]
    assert result["weekly"] == [("2026-10-26", 16, 15, 15 / 16)]
    assert result["monthly"] == [("2026-10", 16, 15, 15 / 16)]
    assert result["overall"] == [("ΟΛΟ ΤΟ ΙΣΤΟΡΙΚΟ", 16, 15, 15 / 16)]
    assert "Alice Secret" not in str(result)
    assert "Bob Secret" not in str(result)
    assert list(MODE_LABELS) == ["daily", "weekly", "monthly", "overall"]


def test_zero_denominator_is_unavailable_not_fake_zero():
    assert department_periods([fact("2026-10-30", "T", 0, 0)])["overall"] == [
        ("ΟΛΟ ΤΟ ΙΣΤΟΡΙΚΟ", 0, 0, None)
    ]


def test_duplicate_or_negative_current_final_fact_fails_closed():
    row = fact("2026-10-30", "T", 2, 2)
    with pytest.raises(ValueError, match="Duplicate"):
        department_periods([row, row])
    with pytest.raises(ValueError, match="Negative"):
        department_periods([fact("2026-10-30", "T", -1, 0)])


def test_export_package_rejects_any_operational_schema_or_vba(tmp_path):
    filename = tmp_path / "public.xlsx"
    with zipfile.ZipFile(filename, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
    inspect_share_package(filename)
    with zipfile.ZipFile(filename, "a") as z:
        z.writestr("xl/vbaProject.bin", "unsafe")
    with pytest.raises(ValueError, match="VBA"):
        inspect_share_package(filename)


def test_export_package_rejects_hidden_operational_schema(tmp_path):
    filename = tmp_path / "public.xlsx"
    with zipfile.ZipFile(filename, "w") as z:
        for i in range(1, 5):
            z.writestr(f"xl/worksheets/sheet{i}.xml", "<worksheet/>")
        z.writestr("xl/workbook.xml", "<workbook>THERAPIST_DAY_HISTORY</workbook>")
    with pytest.raises(ValueError, match="leaked"):
        inspect_share_package(filename)
