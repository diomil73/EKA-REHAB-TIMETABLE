from __future__ import annotations
from datetime import date, time
from pathlib import Path
import json
from rehab_core.provider_policy import LeadershipAssignment, LeadershipRole, OverrideKind, ProviderOverride, ProviderPolicyBook, ProviderProfile, ProviderType

SCHEMA_VERSION = 1

def _d(v): return date.fromisoformat(v) if v else None
def _t(v): return time.fromisoformat(v) if v else None

def load_policy_book(path: str | Path) -> ProviderPolicyBook:
    path = Path(path)
    if not path.exists(): return ProviderPolicyBook()
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != SCHEMA_VERSION:
        raise ValueError("Unsupported provider policy schema")
    book = ProviderPolicyBook()
    for x in payload.get("profiles", []):
        p = ProviderProfile(x["provider_id"], x.get("display_name") or x["provider_id"],
                            ProviderType(x.get("provider_type", "therapist")), bool(x.get("active", True)),
                            bool(x.get("replacement_eligible", True)), x.get("default_max_slots"))
        book.profiles[p.provider_id] = p
    for x in payload.get("overrides", []):
        book.overrides.append(ProviderOverride(x["provider_id"], OverrideKind(x["kind"]), _d(x["start_date"]),
                                              _d(x.get("end_date")), _t(x.get("time_slot")), x.get("max_load"), x.get("reason", "")))
    for x in payload.get("leadership", []):
        book.leadership.append(LeadershipAssignment(x["provider_id"], LeadershipRole(x["role"]),
                                                    _d(x["start_date"]), _d(x.get("end_date")), x.get("note", "")))
    return book

def save_policy_book(book: ProviderPolicyBook, path: str | Path) -> None:
    path = Path(path); path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
      "schema_version": SCHEMA_VERSION,
      "profiles": [{"provider_id": p.provider_id, "display_name": p.display_name,
                    "provider_type": p.provider_type.value, "active": p.active,
                    "replacement_eligible": p.replacement_eligible,
                    "default_max_slots": p.default_max_slots}
                   for p in sorted(book.profiles.values(), key=lambda x: x.provider_id.casefold())],
      "overrides": [{"provider_id": o.provider_id, "kind": o.kind.value,
                     "start_date": o.start_date.isoformat(), "end_date": o.end_date.isoformat() if o.end_date else None,
                     "time_slot": o.time_slot.strftime("%H:%M") if o.time_slot else None,
                     "max_load": o.max_load, "reason": o.reason} for o in book.overrides],
      "leadership": [{"provider_id": l.provider_id, "role": l.role.value,
                      "start_date": l.start_date.isoformat(), "end_date": l.end_date.isoformat() if l.end_date else None,
                      "note": l.note} for l in book.leadership]
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
