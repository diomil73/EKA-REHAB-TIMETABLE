from hashlib import sha256

from openpyxl import Workbook

from tools import audit_integrated_workbook_structure as module


def _book(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "MASTER_SCHEDULE"
    ws.append(["PatientID", "PatientName", "TherapistID", "ServiceDate"])
    ws.append(["SYNTH-P001", "PRIVATE SYNTHETIC PATIENT NAME", "SYNTH-T001", "2026-11-15"])
    ws["E2"] = "=1+1"
    ws["F2"] = "=NormalizeTimeMS(A2)"
    ws["G2"] = "=NormalizePresenceMS(B2)"
    ws["H2"] = '="NormalizeTimeMS(A2)"'
    ws["I2"] = '=A2&"NormalizePresenceMS(B2)"'

    hidden = wb.create_sheet("PATIENT_SESSION_HISTORY")
    hidden.sheet_state = "veryHidden"
    hidden.append(["PatientID", "TherapistID", "SessionStatus"])
    hidden.append(["SYNTH-P001", "SYNTH-T001", "ACTIVE"])

    wb.create_sheet("LOOKUPS")
    wb.save(path)
    wb.close()


def test_structure_audit_is_read_only_and_never_prints_cell_values(tmp_path, capsys):
    path = tmp_path / "synthetic-integrated.xlsm"
    _book(path)
    before = sha256(path.read_bytes()).hexdigest()

    assert module.audit(path) == 0

    output = capsys.readouterr().out
    assert "Worksheet count: 3" in output
    assert "state=veryHidden" in output
    assert "patient_staff_surface=YES" in output
    assert "role_hints=" in output
    assert "formulas=5" in output
    assert "vba_udf_formula_refs=NormalizeTimeMS:1,NormalizePresenceMS:1" in output
    assert "VBA_UDF_FORMULA_REFERENCES: NormalizeTimeMS=1, NormalizePresenceMS=1" in output
    assert "=NormalizeTimeMS(A2)" not in output
    assert "=NormalizePresenceMS(B2)" not in output
    assert "SOURCE HASH UNCHANGED: YES" in output
    assert "PRIVATE SYNTHETIC PATIENT NAME" not in output
    assert "SYNTH-P001" not in output
    assert "SYNTH-T001" not in output
    assert "=1+1" not in output
    assert sha256(path.read_bytes()).hexdigest() == before


def test_structure_audit_rejects_missing_and_unsupported_files(tmp_path, capsys):
    missing = tmp_path / "missing.xlsm"
    assert module.audit(missing) == 2
    assert "SAFETY STOP" in capsys.readouterr().out

    unsupported = tmp_path / "workbook.xlsb"
    unsupported.write_bytes(b"not an excel file")
    assert module.audit(unsupported) == 2
    assert "SAFETY STOP" in capsys.readouterr().out


def test_structure_audit_rejects_corrupt_zip(tmp_path, capsys):
    path = tmp_path / "corrupt.xlsm"
    path.write_bytes(b"not-a-zip")
    assert module.audit(path) == 2
    assert "SAFETY STOP" in capsys.readouterr().out


def test_structure_audit_fails_closed_when_scan_limit_is_reached(tmp_path, monkeypatch, capsys):
    path = tmp_path / "limited.xlsm"
    _book(path)
    monkeypatch.setattr(module, "MAX_CELLS_PER_SHEET", 2)

    assert module.audit(path) == 3

    output = capsys.readouterr().out
    assert "SCAN LIMIT REACHED" in output
    assert "READ-ONLY STRUCTURE AUDIT INCOMPLETE" in output
    assert "SOURCE HASH UNCHANGED: YES" in output



def test_structure_audit_cli_runs_multiple_candidates_and_aggregates_status(monkeypatch, capsys):
    calls = []
    monkeypatch.setattr(module, "audit", lambda path: calls.append(path.name) or 0)

    assert module.main(["first.xlsm", "second.xlsm"]) == 0

    assert calls == ["first.xlsm", "second.xlsm"]
    assert "ALL CANDIDATE STRUCTURE AUDITS COMPLETE" in capsys.readouterr().out


def test_formula_udf_reference_scan_ignores_quoted_string_literals():
    assert module._formula_udf_references("=NormalizeTimeMS(A2)") == {"NormalizeTimeMS"}
    assert module._formula_udf_references("=NormalizePresenceMS(B2)") == {"NormalizePresenceMS"}
    assert module._formula_udf_references('="NormalizeTimeMS(A2)"') == set()
    assert module._formula_udf_references('=A2&"NormalizePresenceMS(B2)"') == set()
    assert module._formula_udf_references("plain text NormalizeTimeMS(A2)") == set()
