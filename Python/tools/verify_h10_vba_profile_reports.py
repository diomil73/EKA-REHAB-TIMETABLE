"""Verify independent, read-only H10 VBA profile evidence without opening Excel.

The input is the sanitized auditor output, not a populated workbook.
Only metadata counts and component fingerprints are processed or printed.
"""

from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path
import re
import sys


EXPECTED = {
    "UNIFIED_H10_DEMO_RECURRING_IDS_20261112_A.xlsm": 67,
    "UNIFIED_H10_DEMO_NATIVE_WRITE_20261112_A.xlsm": 67,
    "UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm": 66,
}

_COMPONENT = re.compile(
    r"^(?:COMPONENT|NO_PATTERN_MATCH_COMPONENT) name=(\S+) "
    r"type=(\d+) lines=(\d+)"
)
_PROFILE = re.compile(
    r"^\s+PROCEDURE_PROFILE name=(\S+) lines=(\d+)-(\d+) "
    r"body_lines=(\d+) categories=(\S+) source_sha256=([0-9a-f]{64})$"
)
_OPERATION = re.compile(
    r"^\s+procedure=(\S+) line=(\d+) operation=([a-z_]+)"
    r"(?: target_sheet=(\S+))?$"
)


_CALL_EDGE = re.compile(
    r"^\s+POSSIBLE_CALL_EDGE caller=(\S+) line=(\d+) callee=(\S+)$"
)


class EvidenceError(ValueError):
    """Incomplete, inconsistent or invalid read-only audit evidence."""


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    return raw.decode("utf-8-sig")


def _parse_block(block: str) -> dict:
    rows = block.splitlines()
    name_line = next((line for line in rows if line.startswith("Workbook: ")), None)
    if name_line is None:
        raise EvidenceError("missing workbook name")
    name = name_line.removeprefix("Workbook: ").strip()
    if name not in EXPECTED:
        raise EvidenceError(f"unexpected candidate: {name}")
    expected_markers = (
        "ReadOnly: True",
        "Macros disabled before open: YES",
        "Excel events disabled: YES",
        "VBA WRITE-SURFACE AUDIT COMPLETE",
        "SOURCE HASH UNCHANGED: YES",
        "NO WORKBOOK WAS MODIFIED",
    )
    for marker in expected_markers:
        if rows.count(marker) != 1:
            raise EvidenceError(f"{name}: missing/duplicate safety marker: {marker}")
    forbidden = ("SAFETY STOP", "AUDIT INCOMPLETE", "OUTPUT TRUNCATED")
    if any(any(bad in line for bad in forbidden) for line in rows):
        raise EvidenceError(f"{name}: incomplete/unsafe audit")

    component_count = f"VBA component count: {EXPECTED[name]}"
    if rows.count(component_count) != 1:
        raise EvidenceError(f"{name}: unexpected VBA component count")

    components: dict[str, tuple] = {}
    operations: list[tuple] = []
    profiles: list[tuple] = []
    call_edges: list[tuple] = []
    component_name: str | None = None
    declared_match_counts: dict[str, int] = {}
    observed_match_counts: Counter[str] = Counter()
    for line in rows:
        match = _COMPONENT.match(line)
        if match:
            component_name = match.group(1)
            if component_name in components:
                raise EvidenceError(f"{name}: duplicate component: {component_name}")
            fingerprint = None
            if line.startswith("NO_PATTERN_MATCH_COMPONENT "):
                match_hash = re.search(r"component_sha256=([0-9a-f]{64})$", line)
                if match_hash is None:
                    raise EvidenceError(f"{name}: missing no-match fingerprint")
                fingerprint = match_hash.group(1)
            components[component_name] = (
                int(match.group(2)),
                int(match.group(3)),
                fingerprint,
            )
            if line.startswith("COMPONENT "):
                matched_count = re.search(r"matched_operations=(\\d+)", line)
                if matched_count is None:
                    raise EvidenceError(
                        f"{name}: missing component matched-operations count"
                    )
                declared_match_counts[component_name] = int(
                    matched_count.group(1)
                )
            continue
        match = _OPERATION.match(line)
        if match:
            if component_name is None:
                raise EvidenceError(f"{name}: operation without component")
            operations.append(
                (component_name, match.group(1), int(match.group(2)),
                 match.group(3), match.group(4) or "")
            )
            observed_match_counts[component_name] += 1
            continue
        match = _CALL_EDGE.match(line)
        if match:
            if component_name is None:
                raise EvidenceError(f"{name}: call edge without component")
            call_edges.append(
                (component_name, match.group(1), int(match.group(2)),
                 match.group(3))
            )
            continue
        match = _PROFILE.match(line)
        if match:
            if component_name is None:
                raise EvidenceError(f"{name}: profile without component")
            profiles.append(
                (component_name, match.group(1), int(match.group(2)),
                 int(match.group(3)), int(match.group(4)), match.group(5),
                 match.group(6))
            )
            continue
        # Unknown audit records are not silently dropped from an
        # otherwise successful comparison.
        if (
            line.lstrip().startswith(
                ("procedure=", "PROCEDURE_PROFILE", "POSSIBLE_CALL_EDGE")
            )
            or line.startswith(("COMPONENT ", "NO_PATTERN_MATCH_COMPONENT "))
        ):
            raise EvidenceError(f"{name}: malformed audit record")

    for recorded_component, expected_count in declared_match_counts.items():
        if observed_match_counts[recorded_component] != expected_count:
            raise EvidenceError(
                f"{name}: component match count mismatch: {recorded_component}"
            )
    total_marker = re.findall(r"^Matched operations: (\d+)$", block, flags=re.MULTILINE)
    if len(total_marker) != 1 or int(total_marker[0]) != len(operations):
        raise EvidenceError(f"{name}: matched-operations count mismatch")
    if not components:
        raise EvidenceError(f"{name}: no VBA components profiled")
    return {
        "name": name,
        "components": components,
        "operations": operations,
        "profiles": profiles,
        "call_edges": call_edges,
        "operation_count": len(operations),
        "categories": Counter(item[3] for item in operations),
    }


def verify_report(path: Path) -> list[dict]:
    text = _read_text(path).replace("\r\n", "\n").replace("\r", "\n")
    blocks = text.split("VBA WRITE-SURFACE READ-ONLY AUDIT\n")
    if len(blocks) != 4 or blocks[0].strip():
        raise EvidenceError("expected exactly three independent audit blocks")
    reports = [_parse_block(block) for block in blocks[1:]]
    if {item["name"] for item in reports} != set(EXPECTED):
        raise EvidenceError("missing or repeated candidate workbooks")
    baseline = reports[0]
    for other in reports[1:]:
        if baseline["components"] != other["components"]:
            raise EvidenceError(f"{other['name']}: nonempty component metadata mismatch")
        if baseline["operations"] != other["operations"]:
            raise EvidenceError(f"{other['name']}: operation metadata mismatch")
        if baseline["profiles"] != other["profiles"]:
            raise EvidenceError(f"{other['name']}: procedure-profile fingerprint mismatch")
        if baseline["call_edges"] != other["call_edges"]:
            raise EvidenceError(f"{other['name']}: possible-call-edge mismatch")
    return reports


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compare three complete, source-redacted H10 VBA audit reports."
    )
    parser.add_argument("report", type=Path)
    args = parser.parse_args()
    try:
        reports = verify_report(args.report)
    except (EvidenceError, OSError, UnicodeError) as exc:
        print(f"SAFETY STOP: {exc}")
        return 2
    baseline = reports[0]
    print("H10 VBA PROFILE EVIDENCE RECONCILIATION GREEN")
    print("Distinct workbooks:", len(reports))
    print("Nonempty component metadata records:", len(baseline["components"]))
    print("Operation-pattern matches per workbook:", baseline["operation_count"])
    print("Procedure-profile hashes:", len(baseline["profiles"]))
    print("Possible call-edge references:", len(baseline["call_edges"]))
    print("Native workbooks untouched: auditor evidence only")
    for category, count in sorted(baseline["categories"].items()):
        print(f"  {category}: {count}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
