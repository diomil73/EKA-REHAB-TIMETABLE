"""Independent read-only check of the proposed PREPARED record after native reopen.

This checker does not write any workbook and never certifies an Apply/Restore.
It is a post-Save/Close/Reopen barrier for later VBA transactional integration.
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
from rehab_excel.daily_input_journal import append_operation_journal, DailyInputJournalError  # noqa: E402
from rehab_excel.daily_input_operation import DailyInputOperation  # noqa: E402
from rehab_excel.daily_input_journal_sheet import read_native_journal_rows  # noqa: E402
from tools.prepare_daily_input_journal_prepared import (  # noqa: E402
    PreparedJournalPreflightError, _existing_journal_rows, _source_date,
    journal_prefix_sha256,
)


class PreparedJournalReadbackError(RuntimeError):
    """Independent PREPARED readback failed, without patient-identifying text."""


def verify_native_excel_readonly(
    workbook_path: str | Path,
    expected_rows: tuple[tuple, ...],
    *,
    excel_factory=None,
) -> str | None:
    """Open the saved workbook in a fresh, macro-disabled native Excel instance.

    Returns the native Excel version when the COM application exposes it.
    """
    if excel_factory is None:
        if sys.platform != "win32":
            raise PreparedJournalReadbackError("Native Excel readback requires Windows")
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError:
            raise PreparedJournalReadbackError("Native Excel COM unavailable") from None
        excel_factory = lambda: win32com.client.DispatchEx("Excel.Application")

    app = None
    workbook = None
    excel_version = None
    try:
        app = excel_factory()
        raw_version = getattr(app, "Version", None)
        if raw_version is not None:
            excel_version = str(raw_version).strip() or None
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        workbook = app.Workbooks.Open(
            str(Path(workbook_path).resolve()), UpdateLinks=0,
            ReadOnly=True, AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if not bool(workbook.ReadOnly):
            raise PreparedJournalReadbackError("Native Excel verification was not read-only")
        worksheet = workbook.Worksheets("DAILY_INPUT_JOURNAL")
        if int(worksheet.Visible) != 2 or not bool(worksheet.ProtectContents):
            raise PreparedJournalReadbackError("Native journal protection or visibility invalid")
        actual_rows = read_native_journal_rows(worksheet)
        if actual_rows != expected_rows:
            raise PreparedJournalReadbackError("Native Excel journal rows differ after reopen")
        return excel_version
    except PreparedJournalReadbackError:
        raise
    except Exception:
        raise PreparedJournalReadbackError("Native Excel read-only reopen failed") from None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def verify_prepared_reopen(
    reopened_xlsm: str | Path,
    handoff: dict[str, object],
    payload_path: str | Path,
    *,
    native_excel: bool = False,
    excel_factory=None,
) -> str | None:
    """Require exactly one proposed PREPARED row appended to unchanged history.

    Returns the native Excel version when native_excel=True and COM exposes it.
    """
    expected_fields = {
        "status", "operation_id", "service_date", "source_sha256",
        "payload_sha256", "existing_journal_rows", "journal_prefix_sha256",
        "prepared_row",
    }
    if type(handoff) is not dict or set(handoff) != expected_fields:
        raise PreparedJournalReadbackError("Prepared handoff schema invalid")
    if handoff["status"] != "PREPARED_NOT_PERSISTED":
        raise PreparedJournalReadbackError("Handoff cannot authorize journal readback")
    if type(handoff["existing_journal_rows"]) is not int or handoff["existing_journal_rows"] < 0:
        raise PreparedJournalReadbackError("Prepared history count invalid")
    for name in ("journal_prefix_sha256", "source_sha256", "payload_sha256"):
        value = handoff[name]
        if (type(value) is not str or len(value) != 64
                or any(ch not in "0123456789abcdef" for ch in value)):
            raise PreparedJournalReadbackError("Prepared digest invalid")
    proposed_row = handoff["prepared_row"]
    if type(proposed_row) is not list or len(proposed_row) != 5:
        raise PreparedJournalReadbackError("Prepared row shape invalid")
    try:
        prepared = DailyInputOperation(
            operation_id=handoff["operation_id"],
            source_sha256=handoff["source_sha256"],
            service_date=handoff["service_date"],
        )
    except (TypeError, ValueError):
        raise PreparedJournalReadbackError("Prepared operation identity invalid") from None

    path = Path(reopened_xlsm).resolve()
    payload = Path(payload_path).resolve()
    if path.suffix.casefold() != ".xlsm" or not path.is_file():
        raise PreparedJournalReadbackError("Reopened workbook unavailable")
    if payload == path or not payload.is_file():
        raise PreparedJournalReadbackError("Prepared payload unavailable")
    if file_sha256(payload) != handoff["payload_sha256"]:
        raise PreparedJournalReadbackError("Prepared payload changed")
    before_hash = file_sha256(path)
    workbook = None
    try:
        workbook = load_workbook(path, read_only=False, data_only=False, keep_vba=True)
        if "DAILY_INPUT" not in workbook.sheetnames:
            raise PreparedJournalReadbackError("Daily sheet unavailable")
        if _source_date(workbook["DAILY_INPUT"]["B2"].value) != prepared.service_date:
            raise PreparedJournalReadbackError("Reopened service date mismatch")
        rows = _existing_journal_rows(workbook)
    except (PreparedJournalReadbackError, PreparedJournalPreflightError) as exc:
        if isinstance(exc, PreparedJournalReadbackError):
            raise
        raise PreparedJournalReadbackError("Reopened journal readback invalid") from None
    except Exception:
        raise PreparedJournalReadbackError("Reopened workbook could not be inspected") from None
    finally:
        if workbook is not None:
            workbook.close()
    if file_sha256(path) != before_hash:
        raise PreparedJournalReadbackError("Reopened workbook changed during readback")
    if file_sha256(payload) != handoff["payload_sha256"]:
        raise PreparedJournalReadbackError("Prepared payload changed during readback")

    prefix_length = handoff["existing_journal_rows"]
    if len(rows) != prefix_length + 1:
        raise PreparedJournalReadbackError("Reopened journal row count mismatch")
    prefix = rows[:prefix_length]
    if journal_prefix_sha256(prefix) != handoff["journal_prefix_sha256"]:
        raise PreparedJournalReadbackError("Historical journal prefix changed")
    try:
        expected = append_operation_journal(prefix, prepared)
    except DailyInputJournalError:
        raise PreparedJournalReadbackError("Reopened operation transition invalid") from None
    if tuple(proposed_row) != expected[-1] or rows != expected:
        raise PreparedJournalReadbackError("Prepared record differs after reopen")
    if native_excel:
        excel_version = verify_native_excel_readonly(path, expected, excel_factory=excel_factory)
        if file_sha256(path) != before_hash:
            raise PreparedJournalReadbackError("Workbook changed during native Excel readback")
        if file_sha256(payload) != handoff["payload_sha256"]:
            raise PreparedJournalReadbackError("Prepared payload changed during native readback")
        return excel_version
    return None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Read-only PREPARED journal Save/Reopen proof.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--handoff", required=True, type=Path)
    parser.add_argument("--payload", required=True, type=Path)
    parser.add_argument(
        "--native-excel", action="store_true",
        help="Also reopen read-only in a fresh native Excel instance with macros disabled.",
    )
    args = parser.parse_args(argv)
    try:
        if len({args.input.resolve(), args.handoff.resolve(), args.payload.resolve()}) != 3:
            raise PreparedJournalReadbackError("Readback paths overlap")
        record = json.loads(args.handoff.read_text(encoding="utf-8"))
        verify_prepared_reopen(
            args.input, record, args.payload, native_excel=args.native_excel,
        )
    except Exception:
        print("SAFETY STOP: DAILY_INPUT PREPARED READBACK FAILED")
        return 2
    print("DAILY_INPUT PREPARED READBACK MATCHED: NO APPLY/RESTORE APPROVAL")
    if args.native_excel:
        print("Native Excel read-only reopen: GREEN")
    print("Workbook unchanged: YES")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
