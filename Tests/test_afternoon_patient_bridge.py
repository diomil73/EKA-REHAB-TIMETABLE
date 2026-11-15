"""Synthetic-only command protocol tests for the existing patient VBA bridge."""
from datetime import time
from pathlib import Path
import os
import subprocess
import sys

from openpyxl import Workbook

BRIDGE = Path(__file__).resolve().parents[1] / "Python" / "tools" / "patient_program_bridge.py"


def _create_synthetic_book(tmp_path, *, afternoon=False, rows=()):
    wb = Workbook()
    patients = wb.active
    patients.title = "PATIENTS"
    patients.append(("PatientID", "Room", "Name", "Infectious", "Status"))
    patients.append(("SYNTH-P1", None, "Synthetic One", None, "ACTIVE"))
    patients.append(("SYNTH-P2", None, "Synthetic Two", None, "ACTIVE"))
    if afternoon:
        pm = wb.create_sheet("AFTERNOON_PHYSIO_SLOTS")
        pm.append(("PatientID", "Treatment", "Weekday", "StartTime", "Therapist"))
        for row in rows:
            pm.append(row)
    output = tmp_path / "SYNTH_BRIDGE.xlsx"
    wb.save(output)
    wb.close()
    return output


def _run(book, patient="SYNTH-P1"):
    return subprocess.run(
        [sys.executable, str(BRIDGE), "afternoon-state-vba", str(book),
         "--patient-id", patient],
        capture_output=True, text=True, encoding="utf-8",
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        check=False,
    )


def test_legacy_workbook_reports_not_installed_without_creating_sheet(tmp_path):
    book = _create_synthetic_book(tmp_path)
    previous = book.read_bytes()
    proc = _run(book)
    assert proc.returncode == 0
    assert proc.stdout.startswith("PM_STATUS\t0\t")
    assert "PM_ROW" not in proc.stdout
    assert book.read_bytes() == previous


def test_header_only_afternoon_sheet_reports_empty(tmp_path):
    book = _create_synthetic_book(tmp_path, afternoon=True)
    proc = _run(book)
    assert proc.returncode == 0
    assert proc.stdout.startswith("PM_STATUS\t1\t")
    assert "κενή" in proc.stdout
    assert "PM_ROW" not in proc.stdout


def test_selected_patient_only_is_returned_as_separate_records(tmp_path):
    book = _create_synthetic_book(tmp_path, afternoon=True, rows=(
        ("SYNTH-P2", "ΦΘ", "ΤΡ", time(17), "SYNTH-PT2"),
        ("SYNTH-P1", "ΦΘ", "ΔΕ", time(16), "SYNTH-PT1"),
    ))
    previous = book.read_bytes()
    proc = _run(book)
    assert proc.returncode == 0
    assert "PM_ROW\tΔΕ\t16:00\tSYNTH-PT1" in proc.stdout
    assert "SYNTH-PT2" not in proc.stdout
    assert "\nTREATMENT\t" not in proc.stdout
    assert "\nCURRENT\t" not in proc.stdout
    assert book.read_bytes() == previous


def test_unknown_patient_fails_closed_without_returning_rows(tmp_path):
    book = _create_synthetic_book(tmp_path, afternoon=True)
    proc = _run(book, "SYNTH-UNKNOWN")
    assert proc.returncode == 2
    assert proc.stdout.startswith("ERROR\t")
    assert "PM_ROW" not in proc.stdout


def test_malformed_afternoon_schema_cannot_be_silently_ignored(tmp_path):
    book = _create_synthetic_book(tmp_path, afternoon=True)
    wb = __import__("openpyxl").load_workbook(book)
    wb["AFTERNOON_PHYSIO_SLOTS"].cell(1, 2).value = "WRONG"
    wb.save(book)
    wb.close()
    proc = _run(book)
    assert proc.returncode == 2
    assert proc.stdout.startswith("ERROR\t")
    assert "PM_ROW" not in proc.stdout


def test_child_python_output_is_utf8_even_if_parent_console_is_greek_cp1253(tmp_path, monkeypatch):
    """Reproduce the Windows decode failure independently of host code page."""
    monkeypatch.setenv("PYTHONIOENCODING", "cp1253")
    book = _create_synthetic_book(tmp_path, afternoon=True, rows=(
        ("SYNTH-P1", "ΦΘ", "ΔΕ", time(16), "SYNTH-PT1"),
    ))
    proc = _run(book)
    assert proc.returncode == 0
    assert "PM_STATUS\t1\t" in proc.stdout
    assert "PM_ROW\tΔΕ\t16:00\tSYNTH-PT1" in proc.stdout
    assert "κενή" not in proc.stdout
