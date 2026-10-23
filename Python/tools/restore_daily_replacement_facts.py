from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.daily_state import build_daily_session_states  # noqa: E402
from rehab_core.daily_statistics import calculate_daily_treatment_statistics  # noqa: E402
from rehab_core.models import ReplacementAssignment  # noqa: E402
from rehab_excel.application_contract import (  # noqa: E402
    ApplicationContractError,
    assert_application_contract,
)
from rehab_excel.daily_input_reader import DailyInputReadError, read_daily_input  # noqa: E402
from rehab_excel.daily_preview_composer import (  # noqa: E402
    DailyPreviewCompositionError,
    compose_outpatient_daily_plan,
)
from rehab_excel.daily_replacement_fact_batch import (  # noqa: E402
    validate_batch_restore_request,
)
from rehab_excel.daily_replacement_facts import (  # noqa: E402
    HEADERS as REPLACEMENT_FACT_HEADERS,
    SHEET_NAME as REPLACEMENT_FACTS_SHEET,
    DailyReplacementFactError,
    _validate_row_against_session,
    read_daily_replacement_fact_rows,
)
from rehab_excel.daily_statistics_writer import (  # noqa: E402
    DailyStatisticsWriteError,
    SHEET_NAME as DAILY_STATISTICS_SHEET,
    build_daily_statistics_rows,
)
from rehab_excel.master_daily_overlay import (  # noqa: E402
    MasterDailyOverlayError,
    build_master_daily_overlay_plan,
    compose_master_daily_overlay,
)
from rehab_excel.master_projection import (  # noqa: E402
    MasterProjectionError,
    build_master_projection,
    remap_master_formula,
)
from rehab_excel.master_projection_writer import (  # noqa: E402
    MasterProjectionWriteError,
    Win32ComMasterProjectionBackend,
)
from rehab_excel.native_excel import (  # noqa: E402
    Win32ComExcelBackend,
    apply_write_plan_to_copy,
)
from rehab_excel.outpatient_presentation import OutpatientPresentationError  # noqa: E402
from rehab_excel.outpatient_schedule_source import (  # noqa: E402
    OutpatientScheduleSourceError,
    read_unified_base_schedule,
)
from rehab_excel.patient_centric_preview import (  # noqa: E402
    PatientCentricPreviewError,
    build_patient_centric_preview_plan,
)
from rehab_excel.patient_registry_source import (  # noqa: E402
    PatientRegistrySourceError,
    read_patient_registry,
)
from tools.restore_daily_replacement_fact import (  # noqa: E402
    AuthoritativeDailyRestoreError,
    LEDGER_SHEET,
    STATUS_RESTORED,
    _has_vba,
    _read_daily_input_date,
    _sha256,
)


def _write_response(path: Path | None, payload: dict[str, object]) -> None:
    """Write the small UTF-8 JSON bridge response when requested by Excel/VBA."""

    if path is None:
        return
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False),
        encoding="utf-8-sig",
    )


def _requested_restore_ids(
    workbook_path: Path,
    *,
    target_date,
    explicit_ids: tuple[str, ...],
    restore_all: bool,
    fact_rows=None,
) -> tuple[str, ...]:
    if restore_all:
        facts = (
            tuple(fact_rows)
            if fact_rows is not None
            else read_daily_replacement_fact_rows(
                workbook_path,
                target_date=target_date,
                include_restored=False,
            )
        )
        ids = tuple(fact.replacement_id for fact in facts if fact.is_active)
        if not ids:
            raise AuthoritativeDailyRestoreError("No ACTIVE replacement facts to restore")
        return ids
    cleaned = tuple(value.strip() for value in explicit_ids if value.strip())
    if not cleaned:
        raise AuthoritativeDailyRestoreError("At least one --replacement-id is required")
    return cleaned


def _active_replacements_after_restore(
    fact_rows,
    *,
    restored_ids: tuple[str, ...],
    sessions,
) -> tuple[ReplacementAssignment, ...]:
    """Derive effective ACTIVE replacements without mutating the staging workbook."""

    restored = set(restored_ids)
    session_by_id = {session.session_id: session for session in sessions}
    assignments: list[ReplacementAssignment] = []
    for fact in fact_rows:
        if not fact.is_active or fact.replacement_id in restored:
            continue
        _validate_row_against_session(fact, session_by_id)
        assignments.append(
            ReplacementAssignment(
                replacement_id=fact.replacement_id,
                target_session_id=fact.target_session_id,
                patient_id=fact.patient_id,
                original_therapist_id=fact.original_therapist_id,
                replacement_therapist_id=fact.replacement_therapist_id,
                replacement_date=fact.replacement_date,
                replacement_time=fact.replacement_time,
                reason=fact.reason,
                replacement_provider_kind=fact.provider_kind,
            )
        )
    return tuple(assignments)


class _RestoreOutputBackend(Win32ComExcelBackend):
    """Apply MASTER rebuild, daily plan, facts, stats and ledger in one Excel session."""

    def __init__(
        self,
        *,
        target_date,
        statistics,
        replacement_ids: tuple[str, ...],
        master_projection,
    ) -> None:
        super().__init__()
        self.target_date = target_date
        self.statistics = statistics
        self.replacement_ids = replacement_ids
        self.master_projection = tuple(master_projection)

    def _apply_master_projection(self, workbook) -> None:
        """Rebuild MASTER with range-level COM writes instead of cell-by-cell traffic."""

        backend = Win32ComMasterProjectionBackend(palette=self.palette)
        try:
            ws = workbook.Worksheets("MASTER_SCHEDULE")
        except Exception as exc:
            raise AuthoritativeDailyRestoreError(
                "Workbook has no MASTER_SCHEDULE sheet"
            ) from exc

        template_row, templates = backend._template_formulas(ws)
        total_cols, _status_col = backend._schema(ws)
        old_last = backend._last_value_row(ws)
        new_last = max(
            template_row - 1,
            max(
                (item.target_row for item in self.master_projection),
                default=template_row - 1,
            ),
        )
        clear_last = max(old_last, new_last)

        if clear_last >= template_row:
            try:
                ws.Range(f"B{template_row}:C{template_row}").Copy()
                ws.Range(f"B{template_row}:C{clear_last}").PasteSpecial(Paste=-4122)
            except Exception:
                pass
            if clear_last > old_last:
                end_col = "M" if total_cols == 13 else "K"
                try:
                    ws.Range(f"A{template_row}:{end_col}{template_row}").Copy()
                    ws.Range(
                        f"A{old_last + 1}:{end_col}{clear_last}"
                    ).PasteSpecial(Paste=-4122)
                except Exception:
                    pass
            ws.Range(f"A{template_row}:P{clear_last}").ClearContents()

        if new_last >= template_row:
            item_by_row = {item.target_row: item for item in self.master_projection}
            formula_matrix: list[tuple[object, ...]] = []
            for row in range(template_row, new_last + 1):
                item = item_by_row.get(row)
                values: list[object] = [None] * total_cols
                if item is not None:
                    for col, template in templates.items():
                        values[col - 1] = remap_master_formula(
                            template,
                            item.planner_row,
                        )
                formula_matrix.append(tuple(values))
            ws.Range(
                ws.Cells(template_row, 1),
                ws.Cells(new_last, total_cols),
            ).Formula = tuple(formula_matrix)

        for item in self.master_projection:
            if item.responsible_doctor:
                backend._render_patient_identity(ws.Cells(item.target_row, 3), item)
            if item.infectious:
                backend._apply_infectious_style(
                    ws.Range(f"B{item.target_row}:C{item.target_row}")
                )

        workbook.Application.CutCopyMode = False

    def _write_statistics(self, workbook) -> None:
        stats_sheet = None
        for sheet in workbook.Worksheets:
            if str(sheet.Name).casefold() == DAILY_STATISTICS_SHEET.casefold():
                stats_sheet = sheet
                break
        if stats_sheet is None:
            stats_sheet = workbook.Worksheets.Add(
                After=workbook.Worksheets(workbook.Worksheets.Count)
            )
            stats_sheet.Name = DAILY_STATISTICS_SHEET
        else:
            stats_sheet.Cells.Clear()

        rows = build_daily_statistics_rows(self.target_date, self.statistics)
        stats_sheet.Range("A1:C1").Merge()
        stats_sheet.Range("A1").Value = "ΗΜΕΡΗΣΙΑ ΣΤΑΤΙΣΤΙΚΑ ΘΕΡΑΠΕΙΩΝ"
        stats_sheet.Range("A1").Font.Bold = True
        stats_sheet.Range("A1").Font.Size = 14
        stats_sheet.Range("A3").Value = "Δείκτης"
        stats_sheet.Range("B3").Value = "Περιγραφή"
        stats_sheet.Range("C3").Value = "Τιμή"
        stats_sheet.Range("A3:C3").Font.Bold = True
        for offset, item in enumerate(rows, start=4):
            stats_sheet.Cells(offset, 1).Value = item.key
            stats_sheet.Cells(offset, 2).Value = item.label
            stats_sheet.Cells(offset, 3).Value = item.value
        stats_sheet.Columns("A").ColumnWidth = 24
        stats_sheet.Columns("B").ColumnWidth = 34
        stats_sheet.Columns("C").ColumnWidth = 18
        try:
            stats_sheet.Range(f"A3:C{3 + len(rows)}").Borders.LineStyle = 1
        except Exception:
            pass

    def _mark_replacement_fact_rows(self, workbook) -> None:
        if not self.replacement_ids:
            return
        try:
            facts = workbook.Worksheets(REPLACEMENT_FACTS_SHEET)
        except Exception as exc:
            raise AuthoritativeDailyRestoreError(
                f"Workbook has no {REPLACEMENT_FACTS_SHEET} sheet"
            ) from exc

        actual_headers = tuple(
            str(facts.Cells(1, col).Value or "").strip()
            for col in range(1, len(REPLACEMENT_FACT_HEADERS) + 1)
        )
        if actual_headers != REPLACEMENT_FACT_HEADERS:
            raise AuthoritativeDailyRestoreError(
                f"{REPLACEMENT_FACTS_SHEET} headers do not match authoritative schema"
            )

        wanted = set(self.replacement_ids)
        last_row = max(int(facts.UsedRange.Rows.Count), 1)
        matches: dict[str, list[int]] = {
            replacement_id: [] for replacement_id in wanted
        }
        for row in range(2, last_row + 1):
            replacement_id = str(facts.Cells(row, 1).Value or "").strip()
            if replacement_id in matches:
                matches[replacement_id].append(row)

        duplicates = [
            replacement_id
            for replacement_id, matched_rows in matches.items()
            if len(matched_rows) > 1
        ]
        missing = [
            replacement_id
            for replacement_id, matched_rows in matches.items()
            if not matched_rows
        ]
        if duplicates:
            raise AuthoritativeDailyRestoreError(
                f"Duplicate replacement fact id(s): {duplicates!r}"
            )
        if missing:
            raise AuthoritativeDailyRestoreError(
                f"Missing replacement fact id(s): {missing!r}"
            )

        for replacement_id in self.replacement_ids:
            row = matches[replacement_id][0]
            current_status = str(facts.Cells(row, 10).Value or "ACTIVE").strip().upper()
            if current_status != "ACTIVE":
                raise AuthoritativeDailyRestoreError(
                    f"Replacement fact {replacement_id!r} is not ACTIVE in output"
                )
            facts.Cells(row, 10).Value = STATUS_RESTORED
        facts.Visible = 2

    def _mark_ledger_rows(self, workbook) -> None:
        if not self.replacement_ids:
            return
        try:
            ledger = workbook.Worksheets(LEDGER_SHEET)
        except Exception:
            return

        wanted = set(self.replacement_ids)
        last_row = int(ledger.Cells(ledger.Rows.Count, 1).End(-4162).Row)
        matches: dict[str, list[int]] = {
            replacement_id: [] for replacement_id in wanted
        }
        for row in range(2, last_row + 1):
            replacement_id = str(ledger.Cells(row, 1).Value or "").strip()
            if replacement_id in matches:
                matches[replacement_id].append(row)

        duplicates = [
            replacement_id
            for replacement_id, matched_rows in matches.items()
            if len(matched_rows) > 1
        ]
        if duplicates:
            raise AuthoritativeDailyRestoreError(
                f"Duplicate ledger ChangeId(s): {duplicates!r}"
            )

        restored_at = datetime.now()
        for replacement_id in self.replacement_ids:
            matched_rows = matches.get(replacement_id, [])
            if not matched_rows:
                continue
            row = matched_rows[0]
            ledger.Cells(row, 12).Value = STATUS_RESTORED
            ledger.Cells(row, 17).Value = restored_at

    def apply(self, workbook_path: Path, patches) -> None:
        if sys.platform != "win32":
            raise AuthoritativeDailyRestoreError(
                "Restore output write-back requires Windows + Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise AuthoritativeDailyRestoreError("pywin32 is required") from exc

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(
                str(workbook_path.resolve()), UpdateLinks=0, ReadOnly=False
            )

            self._apply_master_projection(workbook)
            for patch in patches:
                self._apply_patch(workbook, patch)
            self._write_statistics(workbook)
            self._mark_replacement_fact_rows(workbook)
            self._mark_ledger_rows(workbook)

            try:
                excel.Run(f"'{workbook.Name}'!RefreshDailyChangeLedgerView")
            except Exception:
                pass
            workbook.Save()
        except AuthoritativeDailyRestoreError:
            raise
        except Exception as exc:
            raise AuthoritativeDailyRestoreError(
                f"Could not write authoritative bulk restore output: {exc}"
            ) from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass


def restore_replacements_and_recompute(
    source_path: str | Path,
    output_path: str | Path,
    *,
    replacement_ids: tuple[str, ...] = (),
    restore_all: bool = False,
    overwrite: bool = False,
) -> tuple[int, int, int, bool]:
    """Restore a validated fact set, then rebuild effective daily truth once."""

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise AuthoritativeDailyRestoreError(f"Source workbook not found: {source}")
    if source == output:
        raise AuthoritativeDailyRestoreError("Output must differ from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise AuthoritativeDailyRestoreError("Source and output must both be .xlsm")
    if restore_all and replacement_ids:
        raise AuthoritativeDailyRestoreError(
            "Use either explicit replacement ids or restore_all, not both"
        )
    if output.exists() and not overwrite:
        raise AuthoritativeDailyRestoreError(f"Output already exists: {output}")

    assert_application_contract(source)
    source_before = _sha256(source)
    stage = output.with_name(f".{output.stem}.bulk-restore-stage{output.suffix}")
    if stage.exists():
        stage.unlink()
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, stage)

    try:
        target_date = _read_daily_input_date(stage)
        fact_rows = read_daily_replacement_fact_rows(
            stage,
            target_date=target_date,
            include_restored=True,
        )
        wanted_ids = _requested_restore_ids(
            stage,
            target_date=target_date,
            explicit_ids=replacement_ids,
            restore_all=restore_all,
            fact_rows=fact_rows,
        )
        restored_ids = validate_batch_restore_request(fact_rows, wanted_ids)

        patients = read_patient_registry(stage)
        base_entries = read_unified_base_schedule(stage, patients=patients)
        sessions = materialize_sessions_for_date(base_entries, target_date)
        daily = read_daily_input(stage, patients=patients, sessions=sessions)
        replacements = _active_replacements_after_restore(
            fact_rows,
            restored_ids=restored_ids,
            sessions=sessions,
        )
        states = build_daily_session_states(
            sessions,
            absences=daily.absences,
            replacements=replacements,
            cancellations=daily.cancellations,
            target_date=target_date,
        )
        statistics = calculate_daily_treatment_statistics(states)
        master_projection = build_master_projection(stage, patients=patients)

        preview = build_patient_centric_preview_plan(
            stage,
            base_entries=base_entries,
            states=states,
            sessions=sessions,
            patients=patients,
            rebuild_multi_member_groups=True,
        )
        composed_plan = compose_outpatient_daily_plan(
            preview.write_plan,
            states=states,
            patients=patients,
        )
        master_overlay = build_master_daily_overlay_plan(
            stage,
            base_entries=base_entries,
            states=states,
            patients=patients,
            master_projection_rows=master_projection,
        )
        composed_plan = compose_master_daily_overlay(composed_plan, master_overlay)
        report = apply_write_plan_to_copy(
            composed_plan,
            output,
            backend=_RestoreOutputBackend(
                target_date=target_date,
                statistics=statistics,
                replacement_ids=tuple(restored_ids),
                master_projection=master_projection,
            ),
            overwrite=overwrite,
        )

        source_after = _sha256(source)
        if source_before != source_after:
            raise AuthoritativeDailyRestoreError(
                "Source workbook changed during bulk restore preview"
            )
        vba_preserved = _has_vba(source) == _has_vba(output) and _has_vba(output)
        if not report.source_unchanged or not vba_preserved:
            raise AuthoritativeDailyRestoreError(
                "Bulk restore preview failed source/VBA preservation checks"
            )
        assert_application_contract(output)
        return (
            len(restored_ids),
            len(replacements),
            statistics.unresolved_sessions,
            statistics.is_finalized,
        )
    finally:
        try:
            stage.unlink()
        except OSError:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Restore selected or all ACTIVE daily replacement facts atomically, "
            "then recompute the saved day once."
        )
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--replacement-id", action="append", default=[])
    parser.add_argument("--all-active", action="store_true")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument(
        "--response",
        type=Path,
        help="Optional UTF-8 JSON response file for the Excel/VBA authoritative bridge.",
    )
    args = parser.parse_args()

    input_path = args.input.resolve()
    source_sha256_before = _sha256(input_path) if input_path.exists() else ""

    try:
        restored, active, unresolved, finalized = restore_replacements_and_recompute(
            args.input,
            args.output,
            replacement_ids=tuple(args.replacement_id),
            restore_all=args.all_active,
            overwrite=args.overwrite,
        )
    except (
        ApplicationContractError,
        AuthoritativeDailyRestoreError,
        DailyInputReadError,
        DailyReplacementFactError,
        DailyStatisticsWriteError,
        PatientCentricPreviewError,
        PatientRegistrySourceError,
        OutpatientScheduleSourceError,
        OutpatientPresentationError,
        DailyPreviewCompositionError,
        MasterDailyOverlayError,
        MasterProjectionError,
        MasterProjectionWriteError,
        ValueError,
        KeyError,
    ) as exc:
        _write_response(
            args.response,
            {
                "ok": False,
                "error": str(exc),
                "source_sha256_before": source_sha256_before,
            },
        )
        print(f"SAFETY STOP: {exc}")
        return 2

    output_path = args.output.resolve()
    _write_response(
        args.response,
        {
            "ok": True,
            "output_path": str(output_path),
            "source_sha256_before": source_sha256_before,
            "restored_replacement_facts": restored,
            "active_replacements_after_restore": active,
            "unresolved_sessions_after_recompute": unresolved,
            "finalized_after_recompute": finalized,
            "source_unchanged": True,
            "vba_preserved": True,
            "application_contract_input": True,
            "application_contract_output": True,
        },
    )

    print("AUTHORITATIVE BULK DAILY RESTORE OK")
    print(f"Input: {input_path}")
    print(f"Output: {output_path}")
    print(f"Restored replacement facts: {restored}")
    print(f"Active replacements after restore: {active}")
    print(f"Unresolved sessions after recompute: {unresolved}")
    print(f"Finalized after recompute: {'ΝΑΙ' if finalized else 'ΟΧΙ'}")
    print("Source unchanged: True")
    print("VBA preserved: True")
    print("Application contract input: OK")
    print("Application contract output: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
