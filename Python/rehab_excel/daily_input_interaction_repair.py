from __future__ import annotations

from .daily_input_sheet import THERAPIST_ABSENCE_REASONS


DAILY_INPUT_SHEET = "DAILY_INPUT"
LISTS_SHEET = "_PY_LISTS"
THERAPIST_FIRST_ROW = 7
THERAPIST_LAST_ROW = 16
PATIENT_FIRST_ROW = 21
PATIENT_LAST_ROW = 40
REASON_NAME = "PY_THERAPIST_REASONS"
LEGACY_DIRECT_REPLACEMENT_MACRO = "resolvedailyreplacementsinplace"
DAILY_INPUT_APPLY_MACRO = "ApplyDailyInputPreview"


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
    """Repair native DAILY_INPUT interaction without clearing existing entries."""

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

    ws.Range(f"A{THERAPIST_FIRST_ROW}:F{THERAPIST_LAST_ROW}").Locked = False
    ws.Range(f"A{PATIENT_FIRST_ROW}:E{PATIENT_LAST_ROW}").Locked = False

    rewired = _rewire_legacy_buttons(ws)

    if was_protected:
        ws.Protect(UserInterfaceOnly=True)

    return rewired
