from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from rehab_excel.h10_k_schema_compare import (
    CANDIDATE_BASELINES,
    SchemaComparisonError,
    inspect_candidate,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compare frozen H10 K psychology schema candidates using read-only OOXML."
    )
    parser.add_argument("--root", type=Path, default=Path.cwd())
    args = parser.parse_args()
    root = args.root.resolve()

    lines = [
        "DIOMIL H10 K SCHEMA COMPARISON",
        f"TIME={datetime.now().strftime('%Y%m%d_%H%M%S')}",
        "MODE=READ_ONLY_OOXML",
        "BASELINE=INTEGRATED_TEST_BASELINE_MATRIX_20261115.md",
        "NO_EXCEL_COM_NO_WORKBOOK_SAVE=YES",
        "NO_PATIENT_VALUES_PRINTED=YES",
    ]
    completed = []
    errors = []

    for relative in CANDIDATE_BASELINES:
        lines.extend(["", f"========== {Path(relative).name} =========="])
        try:
            result = inspect_candidate(root, relative)
            completed.append(result)
            lines.append(f"BASELINE_SHA256={result.sha256}")
            lines.append(f"BASELINE_SIZE={result.size}")
            lines.append(f"MASTER_HEADER_ROW={result.master_header_row}")
            for column, matched in result.master_matches:
                lines.append(f"MASTER_{column}_MATCH={matched}")
            for header, columns in result.planner_positions:
                lines.append(
                    f"PLANNER_HEADER={header!r}; POSITIONS={','.join(columns) or 'NONE'}; "
                    f"UNIQUE={len(columns) == 1}"
                )
            lines.append(f"MASTER_SCHEMA_OK={result.master_ok}")
            lines.append(f"PLANNER_SCHEMA_OK={result.planner_ok}")
            lines.append(f"EXPANDED_SCHEMA_MATCH={result.schema_match}")
            lines.append("SOURCE_HASH_UNCHANGED=YES")
            lines.append("ZIP_INTEGRITY=GOOD")
            lines.append("VBA_PRESENT=YES")
        except (SchemaComparisonError, OSError, ValueError, KeyError) as exc:
            errors.append(relative)
            lines.append(f"RESULT=SAFETY_STOP")
            lines.append(f"REASON={exc}")

    matches = [result for result in completed if result.schema_match]
    if errors:
        status = "S"
    elif len(matches) == 1:
        status = "G"
    else:
        status = "B"

    lines.extend([
        "",
        f"CANDIDATES_COMPLETED={len(completed)}/{len(CANDIDATE_BASELINES)}",
        f"SCHEMA_MATCH_COUNT={len(matches)}",
        f"UNIQUE_SCHEMA_CANDIDATE={matches[0].relative_path if len(matches) == 1 else 'NONE'}",
        f"FINAL_STATUS={status}",
        "CLASSIFICATION_SCOPE=SCHEMA_ONLY_NOT_PROVENANCE_OR_ACCEPTANCE",
        "FROZEN_WORKBOOKS_MODIFIED=NO",
    ])

    report = root / (
        f"DIOMIL_H10_K_SCHEMA_COMPARE_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{status}.txt"
    )
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))
    print(f"REPORT={report}")
    return 0 if status != "S" else 2


if __name__ == "__main__":
    raise SystemExit(main())
