"""Read-only weekday/time split proposals for occupational therapy E and EFA H.

No calendar-date inference, patient identifiers or schedule writes.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import re
from dataclasses import dataclass
from pathlib import Path

from tools.parse_real_manual_schedule import TIME_PATTERN, _format_time, _visible_cells

DAYS = {"ΔΕ": "ΔΕ", "ΔΕΥΤ": "ΔΕ", "ΔΕΥ": "ΔΕ", "ΤΡ": "ΤΡ", "ΤΡΙ": "ΤΡ",
        "ΤΕ": "ΤΕ", "ΤΕΤ": "ΤΕ", "ΠΕ": "ΠΕ", "ΠΕΜ": "ΠΕ", "ΠΑ": "ΠΑ", "ΠΑΡ": "ΠΑ"}
DAY_TOKEN = re.compile(
    r"(?<![Α-ΩΆΈΉΊΌΎΏΪΫ])(?:ΔΕΥΤ|ΔΕΥ|ΤΕΤ|ΠΑΡ|ΠΕΜ|ΤΡΙ|ΔΕ|ΤΡ|ΤΕ|ΠΕ|ΠΑ)"
    r"(?![Α-ΩΆΈΉΊΌΎΏΪΫ])"
)
PAREN_DAYS = re.compile(r"\(([^()]*)\)")


@dataclass(frozen=True)
class SplitProposal:
    status: str
    assignments: tuple[tuple[str, str], ...] = ()
    note: str = ""


def _days(fragment: str) -> tuple[str, ...]:
    """Only accept recognized weekday codes and separators, not arbitrary text."""
    remaining = DAY_TOKEN.sub("", fragment.upper())
    if re.sub(r"[\s\-,;/()]+", "", remaining):
        return ()
    return tuple(dict.fromkeys(DAYS[m.group()] for m in DAY_TOKEN.finditer(fragment.upper())))


def _checked_pairs(pairs: list[tuple[str, str]]) -> SplitProposal:
    seen = {}
    for day, time in pairs:
        if day in seen and seen[day] != time:
            return SplitProposal("REVIEW", note=f"Conflicting times for {day}: {seen[day]} vs {time}")
        seen[day] = time
    return SplitProposal("PROPOSED_SPLIT", tuple(seen.items()))


def split_day_times(raw: str, service_column: str = "E") -> SplitProposal:
    """Propose weekday/time pairs, never authorize bookings."""
    cell = str(raw or "").upper().strip()
    # Manual entries may mix Latin E/A into Greek weekday codes (e.g. ΤE, ΠA).
    # This normalized string is for recognition only; source cell is immutable.
    cell = cell.replace("E", "Ε").replace("A", "Α")
    if not cell or cell == "*":
        return SplitProposal("EMPTY")
    times = [(m.start(), m.end(), _format_time(m.group())) for m in TIME_PATTERN.finditer(cell)]
    times = [t for t in times if t[2] is not None]
    if not times:
        return SplitProposal("REVIEW", note="No recognized time")
    if len(times) > 3:
        return SplitProposal("REVIEW", note="Too many time candidates")

    # One time + explicit weekdays anywhere in the cell is a common pattern.
    # E: names of one or more occupational therapists are descriptive, not
    # authoritative assignment IDs. H: require only weekday/time information.
    # Never infer "daily" when E/H omits weekdays.
    if len(times) == 1:
        first = times[0]
        fragment = cell[:first[0]] + " " + cell[first[1]:]
        day_names = tuple(DAYS[m.group()] for m in DAY_TOKEN.finditer(fragment))
        if not day_names:
            return SplitProposal("REVIEW", note="Explicit weekdays absent")
        remainder = DAY_TOKEN.sub(" ", fragment)
        remainder = re.sub(r"\(\s*Τ\s*\)", " ", remainder)  # (Τ) technology annotation
        remainder = re.sub(r"[\s,;()./\-]+", "", remainder)
        if remainder and (service_column != "E" or
                          not re.fullmatch(r"[Α-ΩΆΈΉΊΌΎΏΪΫ]+", remainder)):
            return SplitProposal("REVIEW", note="Unrecognized non-weekday annotation")
        return _checked_pairs([(day, first[2]) for day in day_names])

    # Real E104 format: 11.30(ΔΕ-ΤΕ-ΠΑ)(12.15 ΤΡ-ΠΕ)ΡΑΠΤΗ.
    # Explicit parenthesized day list AFTER first time binds to that time.
    if len(times) == 2:
        first, second = times
        middle = cell[first[1]:second[0]]
        day_group = PAREN_DAYS.search(middle)
        if day_group is not None:
            first_days = _days(day_group.group(1))
            second_days = _days(cell[second[1]:].split(")", 1)[0])
            if first_days and second_days:
                return _checked_pairs(
                    [(day, first[2]) for day in first_days]
                    + [(day, second[2]) for day in second_days]
                )

    # Weekday-first style: e.g. H15, H107. Unparsed names/notes block it.
    pairs = []
    previous_end = 0
    for start, end, clock in times:
        day_group = _days(cell[previous_end:start])
        if not day_group:
            return SplitProposal("REVIEW", note="Day list not uniquely linked to each time")
        pairs.extend((day, clock) for day in day_group)
        previous_end = end
    if cell[previous_end:].strip(" \t,;.)"):
        return SplitProposal("REVIEW", note="Unparsed trailing text or therapist")
    return _checked_pairs(pairs)


def export_review(source: Path, output: Path) -> dict[str, int]:
    """Only write a NEW CSV of proposals; never touch original workbook."""
    if output.exists():
        raise FileExistsError("SAFETY STOP: destination exists; never overwrite")
    if source.resolve() == output.resolve():
        raise ValueError("SAFETY STOP: source/output collision")
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    results, counts = [], {}
    for column in ("E", "H"):
        for cell, raw in _visible_cells(source, column):
            proposal = split_day_times(raw, service_column=column)
            counts[proposal.status] = counts.get(proposal.status, 0) + 1
            for day, clock in proposal.assignments or (("", ""),):
                results.append((cell, column, proposal.status, day, clock, proposal.note))
    with output.open("x", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(("source_cell", "service_column", "status", "weekday", "time", "review_note"))
        writer.writerows(results)
    if hashlib.sha256(source.read_bytes()).hexdigest() != source_hash:
        raise RuntimeError("SAFETY STOP: source changed")
    return counts


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only E/H weekday-time split proposals")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    counts = export_review(args.source, args.output)
    print("E/H WEEKDAY TIME PROPOSALS ONLY; NO IMPORT")
    print("Counts:", counts)
    print("Original source unchanged: True")


if __name__ == "__main__":
    main()
