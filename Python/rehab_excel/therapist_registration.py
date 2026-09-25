from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
import shutil
import sys
import unicodedata
from typing import Protocol

from rehab_core.registration import NewTherapistRequest, validate_new_therapist

from .reader import read_settings


class TherapistRegistrationWriteError(RuntimeError):
    """Raised when a safe therapist-registration preview cannot be produced."""


def _sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _key(value: object) -> str:
    text = unicodedata.normalize("NFD", str(value).strip().casefold())
    return "".join(char for char in text if not unicodedata.combining(char))


def choose_therapist_target_row(last_real_therapist_row: int) -> int:
    """Return the next logical SETTINGS row for THERAPISTS_FTH."""

    return max(2, int(last_real_therapist_row) + 1)


def _remove_preview_file(path: Path, *, required: bool) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        if required:
            raise TherapistRegistrationWriteError(
                "Preview workbook is currently open or locked by Excel. "
                f"Close it and retry: {path}"
            ) from exc


class TherapistRegistrationBackend(Protocol):
    def append_therapist(
        self,
        workbook_path: Path,
        request: NewTherapistRequest,
    ) -> int:
        """Append to SETTINGS!THERAPISTS_FTH and return the Excel row written."""
        ...


class Win32ComTherapistRegistrationBackend:
    """Append one physiotherapist to SETTINGS column A on a preview copy only."""

    @staticmethod
    def _last_used_row(ws, column: int) -> int:
        # xlUp = -4162. This is only an upper bound because formulas/styling can
        # make Excel report cells as used below the real registry values.
        return int(ws.Cells(ws.Rows.Count, column).End(-4162).Row)

    @staticmethod
    def _last_nonblank_value_row(ws, column: int, upper_row: int) -> int:
        upper_row = max(1, int(upper_row))
        values = ws.Range(ws.Cells(1, column), ws.Cells(upper_row, column)).Value2
        if upper_row == 1:
            sequence = ((values,),)
        else:
            sequence = values

        last = 1
        for row_index, item in enumerate(sequence, start=1):
            value = item[0] if isinstance(item, tuple) else item
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            last = row_index
        return last

    def _target_row(self, ws) -> int:
        physical_last = self._last_used_row(ws, 1)
        last_real = self._last_nonblank_value_row(ws, 1, physical_last)
        return choose_therapist_target_row(last_real)

    def append_therapist(
        self,
        workbook_path: Path,
        request: NewTherapistRequest,
    ) -> int:
        if sys.platform != "win32":
            raise TherapistRegistrationWriteError(
                "Therapist preview write-back requires Windows with Microsoft Excel"
            )
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise TherapistRegistrationWriteError(
                "pywin32 is required for therapist preview write-back"
            ) from exc

        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(
                str(workbook_path.resolve()),
                UpdateLinks=0,
                ReadOnly=False,
            )
            try:
                ws = workbook.Worksheets("SETTINGS")
            except Exception as exc:
                raise TherapistRegistrationWriteError("Workbook has no SETTINGS sheet") from exc

            header = str(ws.Cells(1, 1).Value2 or "").strip()
            if header != "THERAPISTS_FTH":
                raise TherapistRegistrationWriteError(
                    "SETTINGS column A is not the expected THERAPISTS_FTH registry"
                )

            target_row = self._target_row(ws)
            if target_row > 2:
                try:
                    ws.Cells(target_row - 1, 1).Copy()
                    ws.Cells(target_row, 1).PasteSpecial(Paste=-4122)  # xlPasteFormats
                except Exception:
                    pass

            ws.Cells(target_row, 1).Value = request.display_name.strip()
            workbook.Application.CutCopyMode = False
            workbook.Save()
            return target_row
        except TherapistRegistrationWriteError:
            raise
        except Exception as exc:
            raise TherapistRegistrationWriteError(
                f"Excel therapist registration failed: {exc}"
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


@dataclass(frozen=True)
class TherapistRegistrationPreviewReport:
    source_path: str
    output_path: str
    therapist_name: str
    excel_row: int
    source_unchanged: bool
    verified_in_output: bool
    other_settings_unchanged: bool


def create_therapist_registration_preview(
    source_path: str | Path,
    output_path: str | Path,
    request: NewTherapistRequest,
    *,
    backend: TherapistRegistrationBackend | None = None,
    overwrite: bool = False,
) -> TherapistRegistrationPreviewReport:
    """Register one therapist in a NEW .xlsm copy and verify SETTINGS values.

    This first write-back stage owns only the physiotherapist name registry in
    SETTINGS column A (THERAPISTS_FTH). Robotic capability is deliberately not
    persisted until an authoritative workbook storage location is defined.
    """

    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.exists():
        raise TherapistRegistrationWriteError(f"Source workbook not found: {source}")
    if source == output:
        raise TherapistRegistrationWriteError("Output must be different from source workbook")
    if source.suffix.casefold() != ".xlsm" or output.suffix.casefold() != ".xlsm":
        raise TherapistRegistrationWriteError("Source and output must both be .xlsm files")
    if output.exists() and not overwrite:
        raise TherapistRegistrationWriteError(f"Output already exists: {output}")
    if request.robotic_capable:
        raise TherapistRegistrationWriteError(
            "Robotic capability has no authoritative therapist storage in SETTINGS yet; "
            "register the therapist without --robotic-capable for now"
        )

    source_settings = read_settings(source)
    check = validate_new_therapist(
        request,
        existing_therapist_names=source_settings.therapist_names,
    )
    if not check.allowed:
        detail = "; ".join(f"{issue.field}: {issue.message}" for issue in check.issues)
        raise TherapistRegistrationWriteError(f"Therapist validation failed: {detail}")

    source_before = _sha256(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        _remove_preview_file(output, required=True)
    shutil.copy2(source, output)

    selected_backend = backend or Win32ComTherapistRegistrationBackend()
    try:
        excel_row = selected_backend.append_therapist(output, request)
    except Exception:
        _remove_preview_file(output, required=False)
        raise

    source_after = _sha256(source)
    if source_before != source_after:
        _remove_preview_file(output, required=False)
        raise TherapistRegistrationWriteError(
            "Source workbook changed during preview creation"
        )

    output_settings = read_settings(output)
    expected_names = source_settings.therapist_names + (request.display_name.strip(),)
    therapist_verified = output_settings.therapist_names == expected_names
    other_settings_unchanged = (
        output_settings.standard_timeslots == source_settings.standard_timeslots
        and output_settings.reclined_timeslots == source_settings.reclined_timeslots
        and output_settings.day_patterns == source_settings.day_patterns
        and output_settings.therapies == source_settings.therapies
        and output_settings.patient_statuses == source_settings.patient_statuses
        and output_settings.yes_no_values == source_settings.yes_no_values
        and output_settings.rooms == source_settings.rooms
        and output_settings.psychologist_names == source_settings.psychologist_names
    )

    normalized_matches = sum(
        1 for name in output_settings.therapist_names if _key(name) == _key(request.display_name)
    )
    verified = therapist_verified and normalized_matches == 1 and other_settings_unchanged
    if not verified:
        _remove_preview_file(output, required=False)
        raise TherapistRegistrationWriteError(
            "Preview verification failed: therapist registry or other SETTINGS values changed unexpectedly"
        )

    return TherapistRegistrationPreviewReport(
        source_path=str(source),
        output_path=str(output),
        therapist_name=request.display_name.strip(),
        excel_row=excel_row,
        source_unchanged=True,
        verified_in_output=True,
        other_settings_unchanged=True,
    )
