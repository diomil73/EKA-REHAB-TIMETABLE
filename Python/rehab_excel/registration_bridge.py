from __future__ import annotations

from dataclasses import asdict
from typing import Callable, Mapping

from .authoritative_commit import file_sha256

from .registration_menu import build_registration_request
from .registration_orchestrator import (
    RegistrationOrchestrationError,
    RegistrationPreviewResult,
    create_registration_preview,
)


class RegistrationBridgeError(RuntimeError):
    """Stable UI-facing error raised by the VBA/Python registration bridge."""


PreviewService = Callable[..., RegistrationPreviewResult]
FingerprintService = Callable[[str], str]


def run_registration_bridge(
    payload: Mapping[str, object],
    *,
    preview_service: PreviewService = create_registration_preview,
    fingerprint_service: FingerprintService = file_sha256,
) -> dict[str, object]:
    """Translate one UI payload into the proven registration preview workflow.

    Expected payload keys:
    - ``source_path``: current authoritative .xlsm path
    - ``preview_dir``: directory where the safe preview copy is created
    - ``action``: new_patient / new_therapist / new_student
    - ``values``: raw form values accepted by ``build_registration_request``
    - ``overwrite``: optional boolean, defaults to False

    The bridge deliberately contains no registration business rules. It only
    adapts UI-shaped data to the existing parser/orchestrator and returns a
    JSON-serialisable result suitable for VBA or another thin client.
    """

    source_path = str(payload.get("source_path", "")).strip()
    preview_dir = str(payload.get("preview_dir", "")).strip()
    action = str(payload.get("action", "")).strip()
    values = payload.get("values")
    overwrite = payload.get("overwrite", False)

    if not source_path:
        raise RegistrationBridgeError("source_path is required")
    if not preview_dir:
        raise RegistrationBridgeError("preview_dir is required")
    if not action:
        raise RegistrationBridgeError("action is required")
    if not isinstance(values, Mapping):
        raise RegistrationBridgeError("values must be an object")
    if not isinstance(overwrite, bool):
        raise RegistrationBridgeError("overwrite must be a boolean")

    try:
        request = build_registration_request(action, values)
        source_sha256_before = fingerprint_service(source_path)
        result = preview_service(
            source_path,
            preview_dir,
            request,
            overwrite=overwrite,
        )
    except (ValueError, TypeError, RegistrationOrchestrationError) as exc:
        raise RegistrationBridgeError(str(exc)) from exc

    response = asdict(result)
    response["kind"] = result.kind.value
    response["source_sha256_before"] = source_sha256_before
    response["action"] = action
    response["ok"] = True
    return response
