from __future__ import annotations

from .daily_input_sheet import THERAPIST_ABSENCE_REASONS


DAILY_INPUT_SHEET = "DAILY_INPUT"
LEDGER_SHEET = "DAILY_CHANGE_LEDGER"
LISTS_SHEET = "_PY_LISTS"
THERAPIST_FIRST_ROW = 7
THERAPIST_LAST_ROW = 16
PATIENT_FIRST_ROW = 21
PATIENT_LAST_ROW = 40
REASON_NAME = "PY_THERAPIST_REASONS"
LEGACY_DIRECT_REPLACEMENT_MACRO = "resolvedailyreplacementsinplace"
DAILY_INPUT_APPLY_MACRO = "ApplyDailyInputPreview"
LEDGER_META_PREFIX = "DAILY_INPUT:"
LEDGER_STATUS_ACTIVE = "ACTIVE"
LEDGER_STATUS_COLUMN = 12
LEDGER_META_COLUMN = 18


class DailyInputInteractionRepairError(RuntimeError):
    pass


def _macro_terminal_name(action: object) -> str:
    text = str(action or "").strip()
    terminal = text.rsplit("!", 1)[-1].strip()
    terminal = terminal.rsplit(".", 1)[-1].strip()
    return terminal.casefold()


def _assert_layout(ws) -> None:
    expected = {
        "A6": "Θεραπευτής",
        "B6": "Όλη ημέρα",
        "E6": "Αιτία",
        "A20": "Ασθενής",
        "D20": "Κατάσταση",
    }
    mismatches = []
    for address, wanted in expected.items():
        actual = str(ws.Range(address).Value or "").strip()
        if actual.casefold() != wanted.casefold():
            mismatches.append(f"{address}: expected {wanted!r}, got {actual!r}")
    if mismatches:
        raise DailyInputInteractionRepairError(
            "Unexpected DAILY_INPUT layout; refusing to repair: " + "; ".join(mismatches)
        )


def _find_sheet(workbook, name: str):
    for sheet in workbook.Worksheets:
        if str(sheet.Name).casefold() == name.casefold():
            return sheet
    return None


def _ensure_reason_list(workbook) -> None:
    lists = _find_sheet(workbook, LISTS_SHEET)
    if lists is None:
        lists = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        lists.Name = LISTS_SHEET

    for row, value in enumerate(THERAPIST_ABSENCE_REASONS, start=1):
        lists.Cells(row, 6).Value = value

    try:
        workbook.Names(REASON_NAME).Delete()
    except Exception:
        pass
    workbook.Names.Add(
        Name=REASON_NAME,
        RefersTo=f"='{LISTS_SHEET}'!$F$1:$F${len(THERAPIST_ABSENCE_REASONS)}",
    )
    lists.Visible = 2


def _set_reason_validation(ws) -> None:
    target = ws.Range(f"E{THERAPIST_FIRST_ROW}:E{THERAPIST_LAST_ROW}")
    try:
        target.Validation.Delete()
    except Exception:
        pass
    target.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=f"={REASON_NAME}")
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ShowError = True


def _parse_daily_input_source_row(meta: object) -> int | None:
    text = str(meta or "").strip()
    if not text.startswith(LEDGER_META_PREFIX):
        return None
    source_row_text = text.rsplit(":", 1)[-1].strip()
    if not source_row_text.isdigit():
        return None
    source_row = int(source_row_text)
    if THERAPIST_FIRST_ROW <= source_row <= THERAPIST_LAST_ROW:
        return source_row
    if PATIENT_FIRST_ROW <= source_row <= PATIENT_LAST_ROW:
        return source_row
    return None


def active_daily_input_rows(workbook) -> tuple[int, ...]:
    """Return source rows whose DAILY_INPUT ledger record is still ACTIVE."""

    ledger = _find_sheet(workbook, LEDGER_SHEET)
    if ledger is None:
        return ()

    last_row = int(ledger.Cells(ledger.Rows.Count, 1).End(-4162).Row)
    active_rows: set[int] = set()
    for row in range(2, last_row + 1):
        status = str(ledger.Cells(row, LEDGER_STATUS_COLUMN).Value or "").strip().upper()
        if status != LEDGER_STATUS_ACTIVE:
            continue
        source_row = _parse_daily_input_source_row(
            ledger.Cells(row, LEDGER_META_COLUMN).Value
        )
        if source_row is not None:
            active_rows.add(source_row)
    return tuple(sorted(active_rows))


def _apply_input_editability(workbook, ws) -> tuple[int, ...]:
    active_rows = active_daily_input_rows(workbook)

    ws.Range(f"A{THERAPIST_FIRST_ROW}:F{THERAPIST_LAST_ROW}").Locked = False
    ws.Range(f"A{PATIENT_FIRST_ROW}:E{PATIENT_LAST_ROW}").Locked = False

    for source_row in active_rows:
        if THERAPIST_FIRST_ROW <= source_row <= THERAPIST_LAST_ROW:
            ws.Range(f"A{source_row}:F{source_row}").Locked = True
        else:
            ws.Range(f"A{source_row}:E{source_row}").Locked = True

    return active_rows


def _rewire_legacy_buttons(ws) -> int:
    rewired = 0

    try:
        buttons = ws.Buttons()
    except Exception:
        buttons = None
    if buttons is not None:
        for index in range(1, int(buttons.Count) + 1):
            button = buttons.Item(index)
            try:
                current = _macro_terminal_name(button.OnAction)
            except Exception:
                current = ""
            if current == LEGACY_DIRECT_REPLACEMENT_MACRO:
                button.OnAction = DAILY_INPUT_APPLY_MACRO
                rewired += 1

    for index in range(1, int(ws.Shapes.Count) + 1):
        shape = ws.Shapes.Item(index)
        try:
            current = _macro_terminal_name(shape.OnAction)
        except Exception:
            current = ""
        if current == LEGACY_DIRECT_REPLACEMENT_MACRO:
            shape.OnAction = DAILY_INPUT_APPLY_MACRO
            rewired += 1

    return rewired


def repair_daily_input_interaction(workbook) -> int:
    """Repair DAILY_INPUT interaction without clearing existing entries.

    Unapplied input rows remain editable. Rows represented by ACTIVE DAILY_INPUT
    ledger records are locked and can only be reversed through THERAPIST_DAILY.
    """

    ws = _find_sheet(workbook, DAILY_INPUT_SHEET)
    if ws is None:
        raise DailyInputInteractionRepairError("Workbook has no DAILY_INPUT sheet")

    _assert_layout(ws)
    was_protected = bool(getattr(ws, "ProtectContents", False))
    if was_protected:
        try:
            ws.Unprotect()
        except Exception as exc:
            raise DailyInputInteractionRepairError(
                "DAILY_INPUT is protected and could not be unlocked safely"
            ) from exc
        if bool(getattr(ws, "ProtectContents", False)):
            raise DailyInputInteractionRepairError(
                "DAILY_INPUT remained protected after unlock attempt"
            )

    _ensure_reason_list(workbook)
    _set_reason_validation(ws)
    _apply_input_editability(workbook, ws)

    rewired = _rewire_legacy_buttons(ws)

    # Protection is required for Locked=True to have native UI effect. Input
    # cells intended for operator editing were explicitly unlocked above.
    ws.Protect(UserInterfaceOnly=True)

    return rewired
