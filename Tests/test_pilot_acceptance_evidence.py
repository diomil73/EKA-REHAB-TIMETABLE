import copy
import json

import pytest

from tools.verify_pilot_acceptance_evidence import (
    PilotEvidenceError, REQUIRED_GATES, main, new_manifest, validate_manifest,
)


def test_fresh_template_cannot_be_mistaken_for_pilot_acceptance():
    manifest = new_manifest()
    pending = validate_manifest(manifest)
    assert all(gate in pending for gate in REQUIRED_GATES)
    assert "native_excel_checked" in pending
    assert "workbook_hash_evidence" in pending


def test_realistic_complete_evidence_requires_all_thirteen_gates_and_native_check():
    manifest = new_manifest()
    manifest["candidate_sha256"] = "a" * 64
    manifest["reference_sha256"] = "b" * 64
    manifest["pilot_start_date"] = "2026-11-12"
    manifest["native_excel_checked"] = True
    for gate in manifest["gates"]:
        gate["status"] = "pass"
        gate["evidence_ref"] = "proof/source_test_01"
    assert validate_manifest(manifest) == ()


def test_claimed_pass_without_evidence_stays_pending():
    manifest = new_manifest()
    manifest["gates"][0]["status"] = "pass"
    assert manifest["gates"][0]["id"] in validate_manifest(manifest)


@pytest.mark.parametrize("mutation", [
    lambda x: x["gates"].pop(),
    lambda x: x["gates"].append(copy.deepcopy(x["gates"][0])),
    lambda x: x["gates"][0].update(status="approved"),
    lambda x: x["gates"][0].update(evidence_ref="patient name here"),
    lambda x: x.update(native_excel_checked="yes"),
    lambda x: x.update(candidate_sha256="nothash"),
    lambda x: x.update(pilot_start_date="2026-13-49"),
])
def test_malformed_or_unsafe_evidence_fails_closed(mutation):
    manifest = new_manifest()
    mutation(manifest)
    with pytest.raises(PilotEvidenceError):
        validate_manifest(manifest)


def test_template_creation_is_non_overwriting_and_remains_pending(tmp_path, monkeypatch):
    from sys import argv
    path = tmp_path / "pilot_evidence.json"
    monkeypatch.setattr("sys.argv", ["tool", "--create-template", str(path)])
    assert main() == 0
    source = path.read_bytes()
    assert all(item["status"] == "pending" for item in json.loads(source)["gates"])
    assert main() == 2
    assert path.read_bytes() == source


def test_manifest_cli_does_not_treat_pending_as_success(tmp_path, monkeypatch):
    path = tmp_path / "pilot_evidence.json"
    path.write_text(json.dumps(new_manifest()), encoding="utf-8")
    monkeypatch.setattr("sys.argv", ["tool", "--manifest", str(path)])
    assert main() == 3


def test_candidate_cannot_be_same_as_reference_workbook():
    manifest = new_manifest()
    manifest["candidate_sha256"] = "a" * 64
    manifest["reference_sha256"] = "a" * 64
    assert "separate_candidate_required" in validate_manifest(manifest)
