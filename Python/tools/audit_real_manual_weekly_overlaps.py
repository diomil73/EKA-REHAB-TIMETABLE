"""Read-only intra-patient weekly slot review for E/H therapist-neutral services.

Flags E/H reservations with the same patient source row, weekday and clock.
These are review hints only: parallel modalities can be legitimate. No writeback.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from tools.parse_real_manual_schedule import _visible_cells
from tools.parse_real_manual_day_time_split import split_day_times


@dataclass(frozen=True)
class DaySlot:
    source_cell: str
    service: str
    row: int
    weekday: str
    time: str


@dataclass(frozen=True)
class ReviewFlag:
    patient_row: int
    weekday: str
    time: str
    source_cells: str
    services: str
    reason: str


def candidates_from_cells(cells: dict[str, str]) -> tuple[list[DaySlot], list[tuple[str, str]]]:
    """Only explicitly supported E/H splits; all others stay excluded and visible."""
    candidates: list[DaySlot] = []
    blocked: list[tuple[str, str]] = []
    for address, raw in sorted(cells.items()):
        match = re.fullmatch(r"([EH])(\d+)", address)
        if not match:
            continue
        service, row = match.group(1), int(match.group(2))
        parsed = split_day_times(raw, service_column=service)
        if parsed.status != "PROPOSED_SPLIT":
            blocked.append((address, parsed.note or parsed.status))
            continue
        for weekday, clock in parsed.assignments:
            candidates.append(DaySlot(address, service, row, weekday, clock))
    return candidates, blocked


def detect_same_patient_weekly_collisions(slots: list[DaySlot]) -> list[ReviewFlag]:
    """Review shared exact times on the SAME source patient row only.

    Same weekday/time across different patients is normal and NOT a collision.
    Neither treatment simultaneity nor shared capacity is decided here.
    """
    groups: dict[tuple[int, str, str], list[DaySlot]] = defaultdict(list)
    for slot in slots:
        groups[(slot.row, slot.weekday, slot.time)].append(slot)
    flags = []
    for (row, weekday, time), matches in sorted(groups.items()):
        sources = sorted({m.source_cell for m in matches})
        services = sorted({m.service for m in matches})
        if len(sources) < 2:
            continue
        flags.append(ReviewFlag(
            row, weekday, time, ";".join(sources), ";".join(services),
            "REVIEW_ONLY: same patient row and weekday/time in multiple services"
        ))
    return flags


def export_review(source: Path, destination: Path) -> dict[str, int]:
    if destination.exists():
        raise FileExistsError("SAFETY STOP: review output already exists")
    if source.resolve() == destination.resolve():
        raise ValueError("SAFETY STOP: cannot write to input")
    original_sha = hashlib.sha256(source.read_bytes()).hexdigest()
    cells = {}
    for col in ("E", "H"):
        cells.update(_visible_cells(source, col))
    candidates, blocked = candidates_from_cells(cells)
    flags = detect_same_patient_weekly_collisions(candidates)
    with destination.open("x", encoding="utf-8-sig", newline="") as fp:
        writer = csv.writer(fp)
        writer.writerow(("kind", "patient_row", "weekday", "time",
                         "source_cells", "services", "reason"))
        for flag in flags:
            writer.writerow(("POSSIBLE_OVERLAP", flag.patient_row, flag.weekday, flag.time,
                             flag.source_cells, flag.services, flag.reason))
        for address, reason in blocked:
            row_number = int(re.search(r"\d+$", address).group())
            writer.writerow(("BLOCKED_SOURCE", row_number, "", "", address, address[0], reason))
    if hashlib.sha256(source.read_bytes()).hexdigest() != original_sha:
        raise RuntimeError("SAFETY STOP: input file hash changed")
    return {
        "candidate_weekday_slots": len(candidates),
        "possible_overlaps": len(flags),
        "blocked_source_cells": len(blocked),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only candidate E/H same-patient overlap review")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    results = export_review(args.source, args.output)
    print("E/H SAME-PATIENT WEEKLY REVIEW ONLY; NO IMPORT")
    print(results)
    print("Source unchanged: True")


if __name__ == "__main__":
    main()
