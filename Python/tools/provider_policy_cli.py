from __future__ import annotations
import argparse, sys
from datetime import date, time
from pathlib import Path
REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON_ROOT = REPO_ROOT / "Python"
if str(PYTHON_ROOT) not in sys.path: sys.path.insert(0, str(PYTHON_ROOT))
from rehab_config.provider_policy_store import load_policy_book, save_policy_book
from rehab_core.provider_policy import LeadershipAssignment, LeadershipRole, OverrideKind, ProviderOverride
DEFAULT_CONFIG = REPO_ROOT / "Config" / "provider_policy.json"

def main() -> int:
    p = argparse.ArgumentParser(description="Manage Diomil provider policy overrides")
    p.add_argument("--config", default=str(DEFAULT_CONFIG)); sub = p.add_subparsers(dest="cmd", required=True)
    s=sub.add_parser("show"); s.add_argument("--date", required=True); s.add_argument("--provider")
    for name in ("set-manager","set-acting-manager"):
        q=sub.add_parser(name); q.add_argument("--provider", required=True); q.add_argument("--from", dest="start", required=True); q.add_argument("--to", dest="end"); q.add_argument("--note", default="")
    q=sub.add_parser("exclude-replacements"); q.add_argument("--provider", required=True); q.add_argument("--from", dest="start", required=True); q.add_argument("--to", dest="end"); q.add_argument("--reason", default="")
    q=sub.add_parser("set-max"); q.add_argument("--provider", required=True); q.add_argument("--max", dest="max_load", type=int, required=True); q.add_argument("--from", dest="start", required=True); q.add_argument("--to", dest="end"); q.add_argument("--reason", default="")
    for name in ("block-time","open-replacement-time"):
        q=sub.add_parser(name); q.add_argument("--provider", required=True); q.add_argument("--time", required=True); q.add_argument("--from", dest="start", required=True); q.add_argument("--to", dest="end"); q.add_argument("--reason", default="")
    a=p.parse_args(); config=Path(a.config); book=load_policy_book(config)
    if a.cmd=="show":
        on=date.fromisoformat(a.date); leader=book.active_leader(on)
        print(f"PROVIDER POLICY | {on}"); print("Active manager: "+(f"{leader.provider_id} ({leader.role.value})" if leader else "not configured"))
        ids=[a.provider] if a.provider else sorted(book.profiles)
        for pid in ids:
            st=book.effective_state(pid,on); b=", ".join(sorted(x.strftime('%H:%M') for x in st.blocked_times)) or "-"; o=", ".join(sorted(x.strftime('%H:%M') for x in st.explicitly_open_replacement_times)) or "-"; role=st.leadership_role.value if st.leadership_role else "normal"
            print(f"{st.display_name} | role {role} | max {st.max_slots} | replacement {'YES' if st.replacement_eligible else 'NO'} | blocked {b} | manager-open {o}")
        return 0
    book.ensure_profile(a.provider, display_name=a.provider)
    if a.cmd in ("set-manager","set-acting-manager"):
        role=LeadershipRole.MANAGER if a.cmd=="set-manager" else LeadershipRole.ACTING_MANAGER
        book.leadership.append(LeadershipAssignment(a.provider, role, date.fromisoformat(a.start), date.fromisoformat(a.end) if a.end else None, a.note))
    else:
        kind={"exclude-replacements":OverrideKind.EXCLUDE_REPLACEMENTS,"set-max":OverrideKind.MAX_LOAD,"block-time":OverrideKind.BLOCK_TIME,"open-replacement-time":OverrideKind.OPEN_REPLACEMENT_TIME}[a.cmd]
        book.overrides.append(ProviderOverride(a.provider, kind, date.fromisoformat(a.start), date.fromisoformat(a.end) if a.end else None,
                                              time.fromisoformat(a.time) if hasattr(a,'time') else None, getattr(a,'max_load',None), getattr(a,'reason','')))
    save_policy_book(book,config); print(f"OK: {config}"); return 0
if __name__=="__main__": raise SystemExit(main())
