from __future__ import annotations

from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Callable, Mapping

from .authoritative_commit import file_sha256

from .outpatient_schedule_registration import (
    OutpatientScheduleRequest,
    OutpatientScheduleWriteError,
    create_outpatient_schedule_preview,
)


class OutpatientScheduleBridgeError(RuntimeError):
    """Stable UI-facing error for outpatient recurring schedule requests."""


def _required_text(values: Mapping[str, object], key: str) -> str:
    value = str(values.get(key, "")).strip()
    if not value:
        raise OutpatientScheduleBridgeError(f"{key} is required")
    return value


def _optional_text(values: Mapping[str, object], key: str) -> str | None:
    value = str(values.get(key, "")).strip()
    return value or None


def _time_value(values: Mapping[str, object], key: str = "time"):
    text = _required_text(values, key)
    try:
        return datetime.strptime(text, "%H:%M").time()
    except ValueError as exc:
        raise OutpatientScheduleBridgeError("time must use HH:MM") from exc


FingerprintService = Callable[[str], str]


def run_outpatient_schedule_bridge(
    payload: Mapping[str, object],
    *,
    fingerprint_service: FingerprintService = file_sha256,
) -> dict[str, object]:
    source_path = str(payload.get("source_path", "")).strip()
    preview_dir = str(payload.get("preview_dir", "")).strip()
    overwrite = payload.get("overwrite", False)
    values = payload.get("values")

    if not source_path:
        raise OutpatientScheduleBridgeError("source_path is required")
    if not preview_dir:
        raise OutpatientScheduleBridgeError("preview_dir is required")
    if not isinstance(overwrite, bool):
        raise OutpatientScheduleBridgeError("overwrite must be a boolean")
    if not isinstance(values, Mapping):
        raise OutpatientScheduleBridgeError("values must be an object")

    request = OutpatientScheduleRequest(
        patient_id=_required_text(values, "patient_id"),
        treatment=_required_text(values, "treatment"),
        start_time=_time_value(values),
        day_pattern=_required_text(values, "days"),
        therapist_id=_optional_text(values, "therapist"),
        target_base_entry_id=_optional_text(values, "target_base_entry_id"),
    )

    output_path = Path(preview_dir) / "OUTPATIENT_SCHEDULE_PREVIEW.xlsm"
    try:
        source_sha256_before = fingerprint_service(source_path)
        report = create_outpatient_schedule_preview(
            source_path,
            output_path,
            request,
            overwrite=overwrite,
        )
    except (OutpatientScheduleWriteError, ValueError) as exc:
        raise OutpatientScheduleBridgeError(str(exc)) from exc

    response = asdict(report)
    response["source_sha256_before"] = source_sha256_before
    response["ok"] = True
    return response
