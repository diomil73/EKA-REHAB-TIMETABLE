"""Read-only non-PREPARED DAILY_INPUT journal transition proposals and checks.

This builds a source-bound proposed next journal row from the existing operation
machine and checks exact append readback. It NEVER authorizes a write. Phase
specific operational evidence has no trusted verifier yet: the output status is
always TRANSITION_PROPOSAL_ONLY, there is no row TSV for VBA, and no mutation,
Apply or Restore action. The proposed row cannot be used as proof that a
business stage completed.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_journal import (  # noqa: E402
    append_operation_journal, decode_operation_journal,
)
from rehab_excel.daily_input_operation import (  # noqa: E402
    DailyInputOperationError, DailyInputOperationPhase as Phase,
    advance_daily_input_operation,
)
from rehab_excel.daily_input_phase_evidence import (  # noqa: E402
    required_evidence_for_transition,
)
from tools.prepare_daily_input_journal_prepared import (  # noqa: E402
    _existing_journal_rows, _source_date, journal_prefix_sha256,
)
from tools.verify_daily_input_journal_prepared import (  # noqa: E402
    verify_native_excel_readonly,
)


class TransitionProposalError(RuntimeError):
    """A non-mutating journal transition proposal or readback failed."""


PROPOSAL_STATUS = "TRANSITION_PROPOSAL_ONLY"
PROPOSAL_FIELDS = frozenset({
    "status", "operation_id", "operation_source_sha256", "service_date",
    "workbook_sha256_before", "journal_prefix_sha256", "existing_journal_rows",
    "previous_phase", "previous_revision", "target_phase", "required_evidence",
    "proposed_row",
})


def _read_source(path: str | Path) -> tuple[str, tuple, str]:
    source = Path(path).resolve()
    if source.suffix.casefold() != ".xlsm" or not source.is_file():
        raise TransitionProposalError("Existing XLSM source required")
    sha_before = file_sha256(source)
    wb = None
    try:
        wb = load_workbook(source, read_only=False, data_only=False, keep_vba=True)
        if "DAILY_INPUT" not in wb.sheetnames:
            raise TransitionProposalError("DAILY_INPUT sheet unavailable")
        service_date = _source_date(wb["DAILY_INPUT"]["B2"].value)
        rows = _existing_journal_rows(wb)
    except TransitionProposalError:
        raise
    except Exception:
        raise TransitionProposalError("Journal source readback invalid") from None
    finally:
        if wb is not None:
            wb.close()
    if file_sha256(source) != sha_before:
        raise TransitionProposalError("Source changed during read-only check")
    return sha_before, rows, service_date


def prepare_transition_proposal(
    source_path: str | Path,
    phase: Phase,
    *,
    operator_resolution: str = "",
    failure_code: str = "",
    failure_detail: str = "",
    rollback_verified: bool = False,
) -> dict[str, object]:
    """Propose, but never authorize, an exact legal next journal revision."""
    if type(phase) is not Phase or phase == Phase.PREPARED:
        raise TransitionProposalError("A non-PREPARED phase is required")
    before_sha, history_rows, workbook_date = _read_source(source_path)
    if not history_rows:
        raise TransitionProposalError("A durable prior journal revision is required")
    previous = decode_operation_journal(history_rows)[-1]
    if previous.service_date != workbook_date:
        raise TransitionProposalError("Journal service date differs from DAILY_INPUT")
    try:
        next_state = advance_daily_input_operation(
            previous, phase, operator_resolution=operator_resolution,
            failure_code=failure_code, failure_detail=failure_detail,
            rollback_verified=rollback_verified,
        )
        required = required_evidence_for_transition(previous, next_state)
        expected_rows = append_operation_journal(history_rows, next_state)
    except (ValueError, DailyInputOperationError):
        raise TransitionProposalError("Illegal journal transition or evidence requirements") from None
    if (len(expected_rows) != len(history_rows) + 1
            or expected_rows[:-1] != history_rows):
        raise TransitionProposalError("Proposed append is not a new revision")
    if file_sha256(source_path) != before_sha:
        raise TransitionProposalError("Workbook changed while preparing proposal")
    return {
        "status": PROPOSAL_STATUS,
        "operation_id": next_state.operation_id,
        "operation_source_sha256": next_state.source_sha256,
        "service_date": next_state.service_date,
        "workbook_sha256_before": before_sha,
        "journal_prefix_sha256": journal_prefix_sha256(history_rows),
        "existing_journal_rows": len(history_rows),
        "previous_phase": previous.phase.value,
        "previous_revision": previous.revision,
        "target_phase": next_state.phase.value,
        "required_evidence": [item.value for item in required],
        "proposed_row": list(expected_rows[-1]),
    }


def _validate_proposal_structure(proposal: dict[str, object]) -> None:
    if type(proposal) is not dict or set(proposal) != PROPOSAL_FIELDS:
        raise TransitionProposalError("Transition proposal schema invalid")
    if proposal["status"] != PROPOSAL_STATUS:
        raise TransitionProposalError("Transition proposal is not read-only")
    for key in ("operation_source_sha256", "workbook_sha256_before", "journal_prefix_sha256"):
        value = proposal[key]
        if (type(value) is not str or len(value) != 64
                or any(char not in "0123456789abcdef" for char in value)):
            raise TransitionProposalError("Transition proposal digest invalid")
    for key in ("operation_id", "service_date", "previous_phase", "target_phase"):
        if type(proposal[key]) is not str or not proposal[key].strip():
            raise TransitionProposalError("Transition proposal identity invalid")
    if (type(proposal["existing_journal_rows"]) is not int
            or proposal["existing_journal_rows"] <= 0
            or type(proposal["previous_revision"]) is not int
            or proposal["previous_revision"] < 0):
        raise TransitionProposalError("Transition proposal revision invalid")
    if type(proposal["required_evidence"]) is not list or not proposal["required_evidence"]:
        raise TransitionProposalError("Transition proof obligations unavailable")
    if (any(type(item) is not str for item in proposal["required_evidence"])
            or len(set(proposal["required_evidence"])) != len(proposal["required_evidence"])):
        raise TransitionProposalError("Transition proof obligations malformed")
    if type(proposal["proposed_row"]) is not list or len(proposal["proposed_row"]) != 5:
        raise TransitionProposalError("Transition row shape invalid")


def verify_transition_source(
    source_path: str | Path,
    proposal: dict[str, object],
) -> None:
    """Check the unchanged preappend source. Does NOT authorize a writer."""
    _validate_proposal_structure(proposal)
    old_sha, rows, service_date = _read_source(source_path)
    if old_sha != proposal["workbook_sha256_before"]:
        raise TransitionProposalError("Workbook changed after transition proposal")
    if service_date != proposal["service_date"]:
        raise TransitionProposalError("Service date changed after proposal")
    if (len(rows) != proposal["existing_journal_rows"]
            or journal_prefix_sha256(rows) != proposal["journal_prefix_sha256"]):
        raise TransitionProposalError("Journal prefix changed after proposal")
    previous = decode_operation_journal(rows)[-1]
    if (previous.operation_id != proposal["operation_id"]
            or previous.source_sha256 != proposal["operation_source_sha256"]
            or previous.phase.value != proposal["previous_phase"]
            or previous.revision != proposal["previous_revision"]):
        raise TransitionProposalError("Previous operation revision differs")
    try:
        target = Phase(proposal["target_phase"])
        candidate_rows = tuple(rows) + (tuple(proposal["proposed_row"]),)
        candidate = decode_operation_journal(candidate_rows)[-1]
        obligations = required_evidence_for_transition(previous, candidate)
    except Exception:
        raise TransitionProposalError("Proposed transition is not legal") from None
    if (candidate.phase != target
            or list(candidate_rows[-1]) != proposal["proposed_row"]
            or [item.value for item in obligations] != proposal["required_evidence"]):
        raise TransitionProposalError("Proposed transition or proof obligations differ")
    if file_sha256(source_path) != old_sha:
        raise TransitionProposalError("Source changed during recheck")


def verify_transition_reopen(
    reopened_path: str | Path,
    proposal: dict[str, object],
    *,
    native_excel: bool = False,
    excel_factory=None,
) -> None:
    """Check exactly one appended revision. Does not prove business effects."""
    _validate_proposal_structure(proposal)
    before_sha, actual_rows, service_date = _read_source(reopened_path)
    prefix_count = proposal["existing_journal_rows"]
    if service_date != proposal["service_date"]:
        raise TransitionProposalError("Reopened service date changed")
    if len(actual_rows) != prefix_count + 1:
        raise TransitionProposalError("Reopened journal revision count differs")
    prefix = actual_rows[:prefix_count]
    if journal_prefix_sha256(prefix) != proposal["journal_prefix_sha256"]:
        raise TransitionProposalError("Reopened journal prefix changed")
    previous, candidate = decode_operation_journal(actual_rows)[-2:]
    if (previous.operation_id != proposal["operation_id"]
            or previous.source_sha256 != proposal["operation_source_sha256"]
            or previous.phase.value != proposal["previous_phase"]
            or previous.revision != proposal["previous_revision"]
            or list(actual_rows[-1]) != proposal["proposed_row"]
            or candidate.phase.value != proposal["target_phase"]):
        raise TransitionProposalError("Reopened transition differs from proposed row")
    try:
        expected_evidence = required_evidence_for_transition(previous, candidate)
    except Exception:
        raise TransitionProposalError("Reopened transition is not legal") from None
    if [item.value for item in expected_evidence] != proposal["required_evidence"]:
        raise TransitionProposalError("Transition proof requirements were changed")
    if native_excel:
        verify_native_excel_readonly(
            reopened_path, actual_rows, excel_factory=excel_factory,
        )
    if file_sha256(reopened_path) != before_sha:
        raise TransitionProposalError("Workbook changed during independent readback")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only non-PREPARED journal transition proposal.")
    parser.add_argument("--input", required=True, type=Path)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--prepare", choices=[p.value for p in Phase if p != Phase.PREPARED])
    mode.add_argument("--verify-source", type=Path)
    mode.add_argument("--verify-reopen", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--operator-resolution", default="")
    parser.add_argument("--failure-code", default="")
    parser.add_argument("--native-excel", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.prepare:
            if (args.output is None or args.native_excel
                    or args.verify_source or args.verify_reopen):
                raise TransitionProposalError("Preparation arguments invalid")
            if (args.output.resolve() == args.input.resolve()
                    or args.output.exists()):
                raise TransitionProposalError("Unused distinct proposal path required")
            result = prepare_transition_proposal(
                args.input, Phase(args.prepare),
                operator_resolution=args.operator_resolution,
                failure_code=args.failure_code,
            )
            args.output.write_text(
                json.dumps(result, ensure_ascii=True, sort_keys=True, separators=(",", ":")),
                encoding="utf-8",
            )
            print("DAILY_INPUT TRANSITION PROPOSAL ONLY: NOT AUTHORIZED FOR WRITE")
        else:
            if args.output is not None or args.operator_resolution or args.failure_code:
                raise TransitionProposalError("Verification arguments invalid")
            proposal_file = args.verify_source or args.verify_reopen
            if proposal_file.resolve() == args.input.resolve():
                raise TransitionProposalError("Proposal path overlaps workbook")
            proposal = json.loads(proposal_file.read_text(encoding="utf-8"))
            if args.verify_source:
                if args.native_excel:
                    raise TransitionProposalError("Native verification is post-reopen only")
                verify_transition_source(args.input, proposal)
                print("DAILY_INPUT TRANSITION SOURCE STILL MATCHES: NOT AUTHORIZED FOR WRITE")
            else:
                verify_transition_reopen(
                    args.input, proposal, native_excel=args.native_excel,
                )
                print("DAILY_INPUT TRANSITION REOPEN MATCHES: NO PHASE APPROVAL")
    except Exception:
        print("SAFETY STOP: DAILY_INPUT TRANSITION PROPOSAL FAILED")
        return 2
    print("Workbook unchanged by this check: YES")
    print("Operational Apply/Restore executed: NO")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
