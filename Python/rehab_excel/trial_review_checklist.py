"""Private, read-only A:V/review task inventory; no clinical decisions.

The new companion XLSX contains only sheet/row/cell locators and generic
reasons, never names, PatientIDs, therapy times or original cell contents.
The existing review remains the authoritative working source. This inventory
is NOT accepted as approval, cannot close a hold, and never edits either input.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from hashlib import sha256
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from rehab_core.programme_mapping import WEEKDAYS, parse_clock
from rehab_excel.trial_admission_gate import (
    CONFIRMED_MARKERS, REQUIRED_SHEETS, SOURCE_HOLDS_SHEET,
    SOURCE_OVERLAP_SHEET, SOURCE_PATIENT_SHEET, SOURCE_WEEKLY_SHEET,
)
from rehab_excel.trial_source_alignment import (
    PROGRAMME_HEADERS, PROGRAMME_SHEET, PATIENT_HEADERS, WEEKLY_HEADERS,
    _headers_match, _row_number,
)

TASK_SHEET = "ΕΚΚΡΕΜΟΤΗΤΕΣ"
TASK_HEADERS = (
    "Κατηγορία", "Πηγή", "Γραμμή/κελιά", "Αιτία",
    "Κατάσταση ελέγχου", "Σημειώσεις χειριστή",
)
PENDING = "ΕΚΚΡΕΜΕΙ"


class TrialReviewChecklistError(ValueError):
    """Generic privacy-safe failure; no clinical values or paths."""


def _digest(path):
    value = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def _text(value):
    return "" if value is None else str(value).strip()


def _rows(sheet, width):
    for number, cells in enumerate(sheet.iter_rows(min_row=2, max_col=width), 2):
        if any(_text(cell.value) for cell in cells):
            yield number, cells


def _unsafe(cells):
    return any(cell.data_type in {"f", "e"} for cell in cells)


def _inside_git(path):
    return any((parent / ".git").exists() for parent in path.parents)


def create_private_review_checklist(*, programme, review, output):
    """Create an exclusive, private companion workbook with pending tasks.

    Snapshot-bound row locators are valid only for these unchanged input
    bytes. Do not edit the original sources to mark tasks complete; use an
    independently authorized review revision and regenerate this inventory.
    """
    p, r, out = (Path(x).resolve() for x in (programme, review, output))
    if (
        len({p, r, out}) != 3
        or any(x.suffix.casefold() != ".xlsx" or _inside_git(x) for x in (p, r, out))
        or not p.is_file() or not r.is_file() or out.exists()
    ):
        raise TrialReviewChecklistError("Distinct private XLSX inputs and a new output outside Git required")
    original = (_digest(p), _digest(r))
    tasks = []
    counts = Counter()

    def task(category, sheet, locator, reason):
        # Only fixed labels and internally generated integer row locators.
        tasks.append((category, sheet, locator, reason, PENDING, None))
        counts[category] += 1

    src = rev = None
    try:
        src = load_workbook(p, read_only=True, data_only=False)
        rev = load_workbook(r, read_only=True, data_only=False)
        if PROGRAMME_SHEET not in src.sheetnames or not REQUIRED_SHEETS.issubset(rev.sheetnames):
            raise TrialReviewChecklistError("Required programme/review sheets missing")
        manual = src[PROGRAMME_SHEET]
        patients = rev[SOURCE_PATIENT_SHEET]
        weekly = rev[SOURCE_WEEKLY_SHEET]
        if (
            not _headers_match(manual, PROGRAMME_HEADERS, start_col=3)
            or not _headers_match(patients, PATIENT_HEADERS)
            or not _headers_match(weekly, WEEKLY_HEADERS)
        ):
            raise TrialReviewChecklistError("Programme/review schema differs")
        source_rows = {}
        for number, cells in _rows(manual, 22):
            source_rows[number] = tuple(cell.value for cell in cells[:2])
            for column, label in ((6, "F_RECLINING"), (21, "U_AFTERNOON"), (22, "V_POOL")):
                if _text(cells[column - 1].value):
                    task(label, PROGRAMME_SHEET, f"{chr(64 + column)}{number}", "SPECIAL_FIELD_REVIEW")
            for index, cell in enumerate(cells, 1):
                if cell.data_type in {"f", "e"}:
                    task("RAW_SOURCE", PROGRAMME_SHEET, f"{chr(64 + index)}{number}", "FORMULA_OR_ERROR")
        identities = {}
        explicit_ids = defaultdict(list)
        for number, cells in _rows(patients, 6):
            source, clinic, name, context, pid, decision = (cell.value for cell in cells)
            row = _row_number(source)
            if (
                _unsafe(cells) or not _text(pid)
                or _text(decision).upper() not in CONFIRMED_MARKERS
            ):
                task("IDENTITY", SOURCE_PATIENT_SHEET, f"E{number}:F{number}", "EXPLICIT_ID_REVIEW")
            else:
                explicit_ids[_text(pid).casefold()].append(number)
            if not _text(context):
                task("CONTEXT", SOURCE_PATIENT_SHEET, f"D{number}", "MISSING_CONTEXT")
            if row is None or row in identities or row not in source_rows:
                task("SOURCE_LINK", SOURCE_PATIENT_SHEET, str(number), "INVALID_OR_REPEATED_SOURCE_ROW")
            else:
                raw_context, raw_name = source_rows[row]
                if (
                    not _text(clinic) or not _text(name)
                    or _text(raw_name) != _text(name) or _text(raw_context) != _text(context)
                ):
                    task("SOURCE_LINK", SOURCE_PATIENT_SHEET, str(number), "SOURCE_IDENTITY_MISMATCH")
                identities[row] = (_text(clinic), _text(name))
        for members in explicit_ids.values():
            if len(members) > 1:
                task("IDENTITY", SOURCE_PATIENT_SHEET, ",".join(map(str, members)), "DUPLICATE_CONFIRMED_ID")
        for sheet_name, width, category in (
            (SOURCE_HOLDS_SHEET, 7, "HOLD"),
            (SOURCE_OVERLAP_SHEET, 7, "REPORTED_OVERLAP"),
        ):
            for number, cells in _rows(rev[sheet_name], width):
                task(category, sheet_name, str(number), "EXISTING_REVIEW_ROW")
        groups = defaultdict(list)
        for number, cells in _rows(weekly, 11):
            clinic, source, name, specialty, day, clock, *_rest, state = (cell.value for cell in cells)
            row = _row_number(source)
            if _text(state) != "PROPOSED_EXACT_SOURCE":
                task("WEEKLY", SOURCE_WEEKLY_SHEET, str(number), "NON_PROPOSED_STATE")
            if _unsafe(cells) or row is None or identities.get(row) != (_text(clinic), _text(name)):
                task("WEEKLY", SOURCE_WEEKLY_SHEET, str(number), "INVALID_SOURCE_LINK_OR_RAW_CELL")
                continue
            try:
                if _text(day) not in WEEKDAYS:
                    raise ValueError("weekday")
                start = parse_clock(clock)
            except (ValueError, TypeError, OverflowError):
                task("WEEKLY", SOURCE_WEEKLY_SHEET, str(number), "INVALID_WEEKDAY_OR_CLOCK")
                continue
            groups[(row, _text(day), start)].append((number, _text(specialty)))
        for members in groups.values():
            if len(members) < 2:
                continue
            category = "CROSS_SPECIALTY" if len({x[1] for x in members}) > 1 else "SAME_SPECIALTY"
            task(category, SOURCE_WEEKLY_SHEET, ",".join(str(x[0]) for x in members), "INDEPENDENT_SLOT_COLLISION")
        # A row omitted from review must not erase original treatment facts.
        for number, cells in _rows(manual, 22):
            if number not in identities and any(_text(cell.value) for cell in cells[2:]):
                task("SOURCE_LINK", PROGRAMME_SHEET, str(number), "TREATMENT_ROW_NOT_IN_REVIEW")
    except TrialReviewChecklistError:
        raise
    except Exception:
        raise TrialReviewChecklistError("Unable to inventory private review safely") from None
    finally:
        if src is not None:
            src.close()
        if rev is not None:
            rev.close()
    if (_digest(p), _digest(r)) != original:
        raise TrialReviewChecklistError("Source changed during checklist preparation")

    wb = Workbook()
    ws = wb.active
    ws.title = TASK_SHEET
    ws.append(TASK_HEADERS)
    for values in tasks:
        ws.append(values)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="22465B")
    for col, width in zip("ABCDEF", (24, 30, 24, 39, 23, 48)):
        ws.column_dimensions[col].width = width
    meta = wb.create_sheet("ΟΔΗΓΙΕΣ")
    meta.append(("Μόνο λίστα ελέγχου. ΟΧΙ έγκριση ή εισαγωγή.",))
    meta.append(("Οι αναφορές ισχύουν μόνο για τα δύο συγκεκριμένα στιγμιότυπα.",))
    meta.append(("Μην αλλάζετε τις πηγές. Χρησιμοποιήστε νέα εγκεκριμένη έκδοση.",))
    meta.append(("SHA256 programme (private)", original[0]))
    meta.append(("SHA256 review (private)", original[1]))
    meta.append(("Δεν περιέχει ονόματα, PatientIDs ή αρχικές κλινικές τιμές.",))
    meta.column_dimensions["A"].width = 80
    meta.column_dimensions["B"].width = 70

    created = False
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("xb") as stream:
            created = True
            wb.save(stream)
        if (_digest(p), _digest(r)) != original:
            raise TrialReviewChecklistError("Source changed during checklist publication")
        check = load_workbook(out, read_only=True, data_only=False)
        try:
            expected = [TASK_HEADERS, *tasks]
            actual = list(check[TASK_SHEET].iter_rows(values_only=True))
            if actual != expected or check["ΟΔΗΓΙΕΣ"]["B4"].value != original[0] or check["ΟΔΗΓΙΕΣ"]["B5"].value != original[1]:
                raise TrialReviewChecklistError("Checklist readback differs")
        finally:
            check.close()
        return dict(counts)
    except Exception:
        if created:
            out.unlink(missing_ok=True)
        raise
    finally:
        wb.close()
