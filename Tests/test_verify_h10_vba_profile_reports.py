"""Synthetic evidence only: do not depend on private or native Excel workbooks."""

from pathlib import Path

import pytest

from tools.verify_h10_vba_profile_reports import EXPECTED, EvidenceError, verify_report


HASH_A = "a" * 64
HASH_B = "b" * 64


def _block(name: str, *, operation="cell_value_or_formula_write",
           fingerprint=HASH_A, complete=True, count=None) -> str:
    count = EXPECTED[name] if count is None else count
    status = ("VBA WRITE-SURFACE AUDIT COMPLETE" if complete
              else "VBA WRITE-SURFACE AUDIT INCOMPLETE: output cap reached")
    return (
        "VBA WRITE-SURFACE READ-ONLY AUDIT\n"
        f"Workbook: {name}\n"
        "ReadOnly: True\n"
        "Macros disabled before open: YES\n"
        "Excel events disabled: YES\n"
        f"VBA component count: {count}\n"
        "NO_PATTERN_MATCH_COMPONENT name=Module5 type=1 lines=1 "
        f"procedures=(none) component_sha256={fingerprint}\n"
        "COMPONENT name=modExample type=1 lines=4 matched_operations=1\n"
        "  procedure=Example line=3 "
        f"operation={operation} target_sheet=DAILY_INPUT\n"
        "Matched operations: 1\n"
        f"{status}\n"
        "SOURCE HASH UNCHANGED: YES\n"
        "NO WORKBOOK WAS MODIFIED\n"
    )


def _write(tmp_path: Path, blocks: list[str], encoding="utf-16") -> Path:
    path = tmp_path / "synthetic_procedures.txt"
    path.write_text("".join(blocks), encoding=encoding)
    return path


def _good_blocks() -> list[str]:
    return [_block(name) for name in EXPECTED]


@pytest.mark.parametrize("encoding", ["utf-8", "utf-16"])
def test_three_synthetic_audits_reconcile(tmp_path, encoding):
    reports = verify_report(_write(tmp_path, _good_blocks(), encoding))
    assert len(reports) == 3
    assert all(x["operation_count"] == 1 for x in reports)
    assert all(len(x["components"]) == 2 for x in reports)


def test_rejects_component_fingerprint_drift(tmp_path):
    blocks = _good_blocks()
    blocks[2] = blocks[2].replace(HASH_A, HASH_B)
    with pytest.raises(EvidenceError, match="component metadata mismatch"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_operation_drift(tmp_path):
    blocks = _good_blocks()
    blocks[1] = blocks[1].replace(
        "cell_value_or_formula_write", "cell_value_or_formula_access"
    )
    with pytest.raises(EvidenceError, match="operation metadata mismatch"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_incomplete_report(tmp_path):
    blocks = _good_blocks()
    blocks[0] = blocks[0].replace(
        "VBA WRITE-SURFACE AUDIT COMPLETE",
        "VBA WRITE-SURFACE AUDIT INCOMPLETE: output cap reached",
    )
    with pytest.raises(EvidenceError, match="missing/duplicate safety marker"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_missing_hash_proof(tmp_path):
    blocks = _good_blocks()
    blocks[2] = blocks[2].replace(
        "SOURCE HASH UNCHANGED: YES", "SOURCE HASH UNCHANGED: NO"
    )
    with pytest.raises(EvidenceError, match="missing/duplicate safety marker"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_duplicate_workbook(tmp_path):
    blocks = _good_blocks()
    blocks[2] = blocks[1]
    with pytest.raises(EvidenceError, match="missing or repeated"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_wrong_component_count(tmp_path):
    blocks = _good_blocks()
    blocks[0] = blocks[0].replace("VBA component count: 67",
                                 "VBA component count: 66")
    with pytest.raises(EvidenceError, match="unexpected VBA component count"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_missing_component_fingerprint(tmp_path):
    blocks = _good_blocks()
    blocks[0] = blocks[0].replace(f"component_sha256={HASH_A}", "")
    with pytest.raises(EvidenceError, match="missing no-match fingerprint"):
        verify_report(_write(tmp_path, blocks))


def test_rejects_tampered_operation_count(tmp_path):
    blocks = _good_blocks()
    blocks[1] = blocks[1].replace("Matched operations: 1",
                                  "Matched operations: 2")
    with pytest.raises(EvidenceError, match="matched-operations count mismatch"):
        verify_report(_write(tmp_path, blocks))



def test_call_edge_metadata_reconciles(tmp_path):
    blocks = _good_blocks()
    for i, block in enumerate(blocks):
        blocks[i] = block.replace(
            "Matched operations: 1",
            "  POSSIBLE_CALL_EDGE caller=Worksheet_Change line=12 "
            "callee=RefreshTherapistDaily\\nMatched operations: 1",
        )
    result = verify_report(_write(tmp_path, blocks))
    assert len(result[0]["call_edges"]) == 1
    assert result[0]["call_edges"][0][-1] == "RefreshTherapistDaily"


def test_call_edge_metadata_drift_is_rejected(tmp_path):
    blocks = _good_blocks()
    for i, block in enumerate(blocks):
        blocks[i] = block.replace(
            "Matched operations: 1",
            "  POSSIBLE_CALL_EDGE caller=Worksheet_Change line=12 "
            "callee=RefreshTherapistDaily\\nMatched operations: 1",
        )
    blocks[2] = blocks[2].replace(
        "callee=RefreshTherapistDaily", "callee=RefreshMaster",
    )
    with pytest.raises(EvidenceError, match="possible-call-edge mismatch"):
        verify_report(_write(tmp_path, blocks))
