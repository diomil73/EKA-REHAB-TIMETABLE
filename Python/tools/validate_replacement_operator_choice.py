from __future__ import annotations

import argparse
from datetime import date, datetime
import json
from pathlib import Path
import sys
import unicodedata

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from rehab_config.provider_policy_store import load_policy_book  # noqa: E402
from rehab_config.service_hours_store import load_service_hours  # noqa: E402
from rehab_core.base_schedule import materialize_sessions_for_date  # noqa: E402
from rehab_core.models import AbsenceKind, Therapist  # noqa: E402
from rehab_core.replacement_policy import find_policy_replacement_options, validate_policy_replacement_time  # noqa: E402
from rehab_core.replacements import create_replacement_assignment  # noqa: E402
from rehab_core.service_replacement_options import build_service_id_by_session, filter_replacement_options_for_service_hours  # noqa: E402
from rehab_core.therapist_absence_queue import build_therapist_absence_replacement_queue  # noqa: E402
from rehab_excel.authoritative_commit import file_sha256  # noqa: E402
from rehab_excel.daily_input_reader import DailyInputReadError, read_daily_input  # noqa: E402
from rehab_excel.daily_replacement_facts import read_daily_replacement_fact_rows, read_daily_replacement_facts  # noqa: E402
from rehab_excel.outpatient_schedule_source import read_unified_base_schedule  # noqa: E402
from rehab_excel.patient_registry_source import read_patient_registry  # noqa: E402
from rehab_excel.reader import read_settings  # noqa: E402
from rehab_excel.student_registry import read_students  # noqa: E402


class ReplacementChoiceValidationError(RuntimeError):
    pass


def _norm(value: str) -> str:
    decomposed = unicodedata.normalize("NFD", str(value).strip().upper())
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def _is_legacy_student_name(name: str) -> bool:
    compact = _norm(name).replace(" ", "")
    return compact.startswith("ΦΟΙΤ") or compact.startswith("FOIT")


def _target_date(path: Path) -> date:
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=True)
    try:
        raw = wb["DAILY_INPUT"]["B2"].value
    finally:
        wb.close()
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    text = str(raw or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise DailyInputReadError(f"Invalid DAILY_INPUT date: {raw!r}")


def _replacement_id(target, history_rows) -> str:
    base = f"daily-input:{target.session_id}:{target.session_date.isoformat()}"
    max_version = 0
    for fact in history_rows:
        replacement_id = str(getattr(fact, "replacement_id", "") or "").strip()
        if replacement_id == base:
            max_version = max(max_version, 1)
            continue
        prefix = f"{base}:v"
        if replacement_id.startswith(prefix):
            suffix = replacement_id[len(prefix):]
            if suffix.isdigit():
                max_version = max(max_version, int(suffix))
    if max_version == 0:
        return base
    return f"{base}:v{max_version + 1}"


def _write_response(path: Path | None, payload: dict[str, object]) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8-sig")


def validate_choice(
    workbook_path: str | Path,
    *,
    target_session_id: str,
    provider_id: str,
    replacement_time: str,
) -> dict[str, object]:
    source = Path(workbook_path).resolve()
    if not source.exists():
        raise ReplacementChoiceValidationError(f"Workbook not found: {source}")
    source_hash = file_sha256(source)

    try:
        chosen_time = datetime.strptime(replacement_time.strip(), "%H:%M").time()
    except ValueError as exc:
        raise ReplacementChoiceValidationError("Replacement time must be HH:MM") from exc

    target_date = _target_date(source)
    patients = tuple(read_patient_registry(source))
    students = tuple(read_students(source))
    settings = read_settings(source)
    base_entries = tuple(read_unified_base_schedule(source))
    sessions = tuple(materialize_sessions_for_date(base_entries, target_date))
    session_by_id = {session.session_id: session for session in sessions}
    target = session_by_id.get(target_session_id)
    if target is None:
        raise ReplacementChoiceValidationError("Target session is not present on the selected date")

    daily = read_daily_input(source, patients=patients, sessions=sessions, therapist_names=settings.therapist_names)
    if not any(item.absence_kind == AbsenceKind.THERAPIST for item in daily.absences):
        raise DailyInputReadError("DAILY_INPUT has no therapist absence rows")

    therapists = tuple(
        Therapist(therapist_id=name, display_name=name)
        for name in settings.therapist_names
        if not _is_legacy_student_name(name)
    )
    policy_book = load_policy_book(REPO_ROOT / "Config" / "provider_policy.json")
    service_hours = load_service_hours(REPO_ROOT / "Config" / "service_hours.json")
    service_id_by_session = build_service_id_by_session(base_entries, target_date)
    existing = tuple(read_daily_replacement_facts(source, sessions=sessions, target_date=target_date))
    queue = build_therapist_absence_replacement_queue(
        sessions=sessions, absences=daily.absences, cancellations=daily.cancellations,
        replacements=existing, therapists=therapists, patients=patients,
        timeslots=settings.standard_timeslots, students=students, policy_book=policy_book,
        base_entries=base_entries,
    )
    unresolved_ids = {item.session_id for item in queue.items}
    if target_session_id not in unresolved_ids:
        raise ReplacementChoiceValidationError("Target session is no longer unresolved; refresh the replacement screen")

    options = find_policy_replacement_options(
        target_session=target, therapists=therapists, sessions=sessions, policy_book=policy_book,
        absences=daily.absences, patients=patients, replacements=existing,
        cancellations=daily.cancellations, requested_time=target.start_time,
        timeslots=settings.standard_timeslots, students=students, base_entries=base_entries,
    )
    options = filter_replacement_options_for_service_hours(
        options, target_session=target, hours_book=service_hours,
        explicit_service_id=service_id_by_session.get(target.session_id),
    )
    chosen_option = next(
        (option for option in options if option.provider_id == provider_id and chosen_time in option.available_timeslots),
        None,
    )
    if chosen_option is None:
        raise ReplacementChoiceValidationError("The selected therapist/time is no longer a valid option")

    validate_policy_replacement_time(
        provider_id=provider_id, target_date=target.session_date,
        replacement_time=chosen_time, policy_book=policy_book,
    )
    history = tuple(read_daily_replacement_fact_rows(source, target_date=target_date, include_restored=True))
    replacement = create_replacement_assignment(
        replacement_id=_replacement_id(target, history),
        target_session=target,
        replacement_therapist_id=provider_id,
        therapists=therapists,
        sessions=sessions,
        absences=daily.absences,
        replacements=existing,
        cancellations=daily.cancellations,
        replacement_time=chosen_time,
        reason="DAILY_INPUT therapist absence",
        patients=patients,
    )

    if file_sha256(source) != source_hash:
        raise ReplacementChoiceValidationError("Workbook changed during replacement choice validation")

    return {
        "ok": True,
        "replacement_id": replacement.replacement_id,
        "target_session_id": replacement.target_session_id,
        "patient_id": replacement.patient_id,
        "original_therapist_id": replacement.original_therapist_id,
        "replacement_therapist_id": replacement.replacement_therapist_id,
        "replacement_date": replacement.replacement_date.isoformat(),
        "replacement_time": replacement.replacement_time.strftime("%H:%M"),
        "reason": replacement.reason,
        "provider_kind": replacement.replacement_provider_kind.value,
        "status": "ACTIVE",
        "source_unchanged": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate one operator replacement choice.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--target-session-id", required=True)
    parser.add_argument("--provider-id", required=True)
    parser.add_argument("--time", required=True)
    parser.add_argument("--response", required=True, type=Path)
    args = parser.parse_args()
    try:
        result = validate_choice(
            args.input,
            target_session_id=args.target_session_id,
            provider_id=args.provider_id,
            replacement_time=args.time,
        )
    except Exception as exc:
        _write_response(args.response, {"ok": False, "error": str(exc)})
        print(f"SAFETY STOP: {exc}")
        return 2
    _write_response(args.response, result)
    print("REPLACEMENT OPERATOR CHOICE VERIFIED")
    print("Input unchanged: True")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
