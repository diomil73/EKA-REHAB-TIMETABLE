"""Private real-manual-source DIOMIL practice DEMO entry point.

All original XLSX and clinical approvals remain on an authorized local Windows
disk, never in the Git checkout or CI. Default mode is READ-ONLY preflight.
Only --build with an exclusive new private --output can write native Excel.
The CLI never logs clinical values, paths, hospital IDs or source exception text.
"""
from __future__ import annotations

import argparse
import json
from hashlib import sha256
from pathlib import Path
import sys
from typing import Mapping

ROOT = Path(__file__).resolve().parents[2]
PYTHON = ROOT / "Python"
if str(PYTHON) not in sys.path:
    sys.path.insert(0, str(PYTHON))

from rehab_core.models import PatientType  # noqa: E402
from rehab_excel.trial_demo_native_seed import (  # noqa: E402
    assert_clean_native_demo_shell, seed_trial_demo_workbook,
)
from rehab_excel.trial_demo_source_projection import (  # noqa: E402
    project_trial_demo_sources, merge_trial_demo_outpatients,
)
from rehab_excel.trial_private_outpatient_projection import (  # noqa: E402
    project_private_outpatient_source,
)
from rehab_excel.trial_provider_schedule import (  # noqa: E402
    read_private_provider_schedule, merge_with_approved_provider_schedule,
)
from rehab_excel.trial_import_plan import (  # noqa: E402
    prepare_trial_import_plan,
)
from rehab_excel.trial_patient_registry_projection import (  # noqa: E402
    ConfirmedClinicalFields,
)
from rehab_excel.trial_special_fields_audit import (  # noqa: E402
    inspect_real_manual_special_fields, verify_legacy_afternoon_scope,
)


APPROVAL_SCHEMA = "diomil-private-clinical-approvals-v1"
REQUIRED_FIELDS = frozenset({
    "clinic", "source_row", "patient_type", "infectious",
})
OPTIONAL_FIELDS = frozenset({
    "room", "responsible_doctor", "hospital_mrn", "status",
})
ALLOWED_FIELDS = REQUIRED_FIELDS | OPTIONAL_FIELDS


class RealDemoCommandError(ValueError):
    """Invalid or unsafe private inputs. Never include real source values."""


def _object_unique(items):
    result = {}
    for key, value in items:
        if key in result:
            raise RealDemoCommandError("Duplicate JSON approval field")
        result[key] = value
    return result


def _text_or_none(value: object) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value.strip():
        raise RealDemoCommandError("Optional clinical field must be text or null")
    return value.strip()


def read_private_clinical_approvals(
    path: str | Path,
) -> Mapping[tuple[str, int], ConfirmedClinicalFields]:
    """Strict JSON schema; does NOT infer clinical facts from coloured cells."""
    source = Path(path)
    if source.suffix.lower() != ".json" or not source.is_file():
        raise RealDemoCommandError("Private approval JSON is required")
    try:
        with source.open("r", encoding="utf-8") as handle:
            document = json.load(handle, object_pairs_hook=_object_unique)
    except (OSError, UnicodeError, json.JSONDecodeError):
        raise RealDemoCommandError("Private approval JSON cannot be parsed") from None
    if (
        not isinstance(document, dict) or set(document) != {"schema", "records"}
        or document["schema"] != APPROVAL_SCHEMA
        or not isinstance(document["records"], list)
        or not document["records"]
    ):
        raise RealDemoCommandError("Private approval root schema invalid")
    result: dict[tuple[str, int], ConfirmedClinicalFields] = {}
    seen_casefold: set[tuple[str, int]] = set()
    for item in document["records"]:
        if (
            not isinstance(item, dict)
            or not REQUIRED_FIELDS.issubset(item)
            or not set(item).issubset(ALLOWED_FIELDS)
        ):
            raise RealDemoCommandError("Clinical approval record schema invalid")
        clinic, row, kind, infectious = (
            item["clinic"], item["source_row"],
            item["patient_type"], item["infectious"],
        )
        if (
            not isinstance(clinic, str) or not clinic.strip()
            or clinic != clinic.strip()
            or type(row) is not int or row < 2
            or type(infectious) is not bool
            or kind not in ("INPATIENT", "OUTPATIENT")
        ):
            raise RealDemoCommandError("Clinical approval classification invalid")
        key = (clinic, row)
        normalized = (clinic.casefold(), row)
        if key in result or normalized in seen_casefold:
            raise RealDemoCommandError("Duplicate clinical source-row approval")
        seen_casefold.add(normalized)
        result[key] = ConfirmedClinicalFields(
            patient_type=PatientType[kind],
            infectious=infectious,
            room=_text_or_none(item.get("room")),
            responsible_doctor=_text_or_none(item.get("responsible_doctor")),
            hospital_mrn=_text_or_none(item.get("hospital_mrn")),
            status=_text_or_none(item.get("status")),
        )
    return result


def _approval_digest(path: Path) -> str:
    """Bind private clinical approval bytes to this invocation, no JSON output."""
    digest = sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _inside_git_checkout(path: Path) -> bool:
    resolved = path.resolve()
    return any((parent / ".git").exists() for parent in (
        resolved.parent, *resolved.parent.parents,
    ))


def validate_private_paths(
    *, programme: Path, review: Path, shell: Path,
    approvals: Path, output: Path | None, build: bool,
    outpatients: Path | None = None,
    provider_schedule: Path | None = None,
) -> None:
    """Ensure neither clinical sources nor final output occupy a Git tree."""
    if build and outpatients is None:
        raise RealDemoCommandError("Full native real-data demo needs approved outpatients")
    sources = (programme, review, shell, approvals) + (
        (outpatients,) if outpatients is not None else ()
    )
    extensions = (".xlsx", ".xlsx", ".xlsm", ".json") + (
        (".xlsx",) if outpatients is not None else ()
    )
    if provider_schedule is not None:
        if outpatients is None:
            raise RealDemoCommandError("Provider schedule requires separate outpatient source")
        sources += (provider_schedule,)
        extensions += (".json",)
    if any(
        not source.is_file() or source.suffix.casefold() != extension
        or _inside_git_checkout(source)
        for source, extension in zip(sources, extensions)
    ):
        raise RealDemoCommandError("Private source files must be outside Git")
    if len({source.resolve() for source in sources}) != len(sources):
        raise RealDemoCommandError("Private source files must be distinct")
    if not build:
        if output is not None:
            raise RealDemoCommandError("Output is only allowed with --build")
        return
    if output is None:
        raise RealDemoCommandError("Explicit output required for native build")
    if sys.platform != "win32":
        raise RealDemoCommandError("Native real-data build requires Windows Excel")
    destination = output.resolve()
    if (
        output.suffix.casefold() != ".xlsm"
        or "_DEMO_" not in output.name.upper()
        or output.exists()
        or _inside_git_checkout(output)
        or destination in {source.resolve() for source in sources}
    ):
        raise RealDemoCommandError("Unsafe private DEMO destination")


def run_private_trial(
    *, programme: Path, review: Path, shell: Path,
    approvals: Path, output: Path | None = None,
    build: bool = False,
    outpatients: Path | None = None,
    provider_schedule: Path | None = None,
) -> tuple[int, int, bool]:
    """Return non-identifying patient/weekday-slot counts and native flag."""
    validate_private_paths(
        programme=programme, review=review, shell=shell,
        approvals=approvals, output=output, build=build,
        outpatients=outpatients, provider_schedule=provider_schedule,
    )
    # Mandatory before any clinical projection or native Excel write:
    # raw formula-only F/U/V cells must never pass data_only source parsing.
    special_fields = inspect_real_manual_special_fields(programme)
    if special_fields.orphan_rows or (
        special_fields.unresolved_fields
        and (special_fields.reclining_rows or special_fields.pool_rows
             or special_fields.formula_fields)
    ):
        raise RealDemoCommandError(
            "Special F/U/V source fields require independent approval"
        )
    if special_fields.unresolved_fields:
        try:
            verify_legacy_afternoon_scope(programme=programme, review=review)
        except ValueError:
            raise RealDemoCommandError(
                "Historical U cells require explicit non-import acknowledgment"
            ) from None
    if build and provider_schedule is None:
        raise RealDemoCommandError("Full native DEMO requires approved provider schedule")
    clinical_before = _approval_digest(approvals)
    schedule_before = _approval_digest(provider_schedule) if provider_schedule is not None else None
    schedule = read_private_provider_schedule(provider_schedule) if provider_schedule is not None else None
    clinical = read_private_clinical_approvals(approvals)
    assert_clean_native_demo_shell(shell)
    plan = prepare_trial_import_plan(
        review=review, intake=shell, programme=programme,
        empty_demo_shell=True,
    )
    projection = project_trial_demo_sources(
        plan, programme=programme, review=review, intake=shell,
        clinical_approvals=clinical,
    )
    if outpatients is not None:
        separate_outpatients = project_private_outpatient_source(outpatients)
        if schedule is None:
            projection = merge_trial_demo_outpatients(projection, separate_outpatients)
        else:
            projection = merge_with_approved_provider_schedule(
                projection, separate_outpatients, plan=plan, schedule=schedule,
            )
    # No use of partial fixture rows: entire original A:V source is checked
    # through the production reconciliation gate above.
    plan.verify_sources(review, shell, programme=programme)
    if _approval_digest(approvals) != clinical_before:
        raise RealDemoCommandError("Private clinical approvals changed during preflight")
    if provider_schedule is not None and _approval_digest(provider_schedule) != schedule_before:
        raise RealDemoCommandError("Private provider schedule changed during preflight")
    if not build:
        return projection.patients, projection.weekday_slot_count, False
    # No clinical values or local path are printed, even after success.
    assert output is not None
    report = seed_trial_demo_workbook(
        plan, programme=programme, review=review, intake=shell,
        output=output, clinical_approvals=clinical,
        outpatient_source=outpatients, provider_schedule=schedule,
    )
    if _approval_digest(approvals) != clinical_before:
        # The sealed in-memory data and local JSON no longer agree. Preserve
        # the candidate for safe private review but DO NOT declare it accepted.
        raise RealDemoCommandError("Clinical approvals changed during native build")
    if provider_schedule is not None and _approval_digest(provider_schedule) != schedule_before:
        raise RealDemoCommandError("Private provider schedule changed during native build")
    if not report.verified_native_readback or not report.source_unchanged:
        raise RealDemoCommandError("Native DEMO verification incomplete")
    return report.patient_count, report.weekday_slot_count, True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Private actual-source Diomil read-only preflight / native build"
    )
    parser.add_argument("--programme", required=True, type=Path)
    parser.add_argument("--review", required=True, type=Path)
    parser.add_argument("--shell", required=True, type=Path)
    parser.add_argument("--clinical-approvals", required=True, type=Path)
    parser.add_argument("--outpatients", type=Path)
    parser.add_argument("--provider-schedule", type=Path)
    parser.add_argument("--build", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    try:
        patients, slots, built = run_private_trial(
            programme=args.programme, review=args.review,
            shell=args.shell, approvals=args.clinical_approvals,
            output=args.output, build=args.build,
            outpatients=args.outpatients, provider_schedule=args.provider_schedule,
        )
    except Exception as exc:
        # No traceback, exception message, source name, patient name, ID,
        # doctor or private file path can leak into public diagnostics.
        print(f"SAFETY STOP: {type(exc).__name__}; no approved native result.")
        return 2
    print("DIOMIL PRIVATE REAL-SOURCE " + (
        "NATIVE DEMO CANDIDATE VERIFIED" if built
        else ("READ-ONLY PREFLIGHT GREEN" if args.provider_schedule is not None
              else "PARTIAL READ-ONLY SOURCE CHECK; FULL PREFLIGHT PENDING")
    ))
    print(f"Confirmed patients: {patients}")
    print(f"Exact Monday-Friday treatment placements: {slots}")
    if args.outpatients is None:
        print("Separate outpatient source: NOT PROVIDED; full DEMO is NOT eligible.")
    else:
        print("Separate outpatient source: AUDITED and included in source counts.")
    if args.provider_schedule is None:
        print("Provider schedule: NOT PROVIDED; full DEMO is NOT eligible.")
    else:
        print("Provider schedule: source-bound named clock/capacity and shared membership checked.")
    print("Source XLSX/review/clinical approvals and original XLSM unchanged: YES")
    if built:
        print("New isolated private DEMO created. OPERATOR ACCEPTANCE PENDING.")
    else:
        print("NO XLSM WRITTEN. Use --build only after private approval.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
