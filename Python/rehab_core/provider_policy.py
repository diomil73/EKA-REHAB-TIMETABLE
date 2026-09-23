from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, time
from enum import Enum
from typing import Iterable


class ProviderType(str, Enum):
    THERAPIST = "therapist"
    STUDENT = "student"


class LeadershipRole(str, Enum):
    MANAGER = "manager"
    ACTING_MANAGER = "acting_manager"


class OverrideKind(str, Enum):
    EXCLUDE_REPLACEMENTS = "exclude_replacements"
    MAX_LOAD = "max_load"
    BLOCK_TIME = "block_time"
    OPEN_REPLACEMENT_TIME = "open_replacement_time"


@dataclass(frozen=True)
class ProviderProfile:
    provider_id: str
    display_name: str
    provider_type: ProviderType = ProviderType.THERAPIST
    active: bool = True
    replacement_eligible: bool = True
    default_max_slots: int | None = None

    @property
    def normal_max_slots(self) -> int:
        if self.default_max_slots is not None:
            return self.default_max_slots
        return 5 if self.provider_type == ProviderType.STUDENT else 6


@dataclass(frozen=True)
class ProviderOverride:
    provider_id: str
    kind: OverrideKind
    start_date: date
    end_date: date | None = None
    time_slot: time | None = None
    max_load: int | None = None
    reason: str = ""

    def is_active(self, on_date: date) -> bool:
        return on_date >= self.start_date and (self.end_date is None or on_date <= self.end_date)


@dataclass(frozen=True)
class LeadershipAssignment:
    provider_id: str
    role: LeadershipRole
    start_date: date
    end_date: date | None = None
    note: str = ""

    def is_active(self, on_date: date) -> bool:
        return on_date >= self.start_date and (self.end_date is None or on_date <= self.end_date)


@dataclass(frozen=True)
class EffectiveProviderState:
    provider_id: str
    display_name: str
    active: bool
    provider_type: ProviderType
    leadership_role: LeadershipRole | None
    max_slots: int
    replacement_eligible: bool
    blocked_times: frozenset[time] = frozenset()
    explicitly_open_replacement_times: frozenset[time] = frozenset()
    applied_override_reasons: tuple[str, ...] = ()

    @property
    def is_active_leader(self) -> bool:
        return self.leadership_role in {LeadershipRole.MANAGER, LeadershipRole.ACTING_MANAGER}


@dataclass(frozen=True)
class PolicyDecision:
    allowed: bool
    reason: str
    effective_state: EffectiveProviderState


@dataclass
class ProviderPolicyBook:
    profiles: dict[str, ProviderProfile] = field(default_factory=dict)
    overrides: list[ProviderOverride] = field(default_factory=list)
    leadership: list[LeadershipAssignment] = field(default_factory=list)

    def ensure_profile(self, provider_id: str, *, display_name: str | None = None,
                       provider_type: ProviderType = ProviderType.THERAPIST) -> ProviderProfile:
        profile = self.profiles.get(provider_id)
        if profile is not None:
            return profile
        profile = ProviderProfile(provider_id=provider_id, display_name=display_name or provider_id,
                                  provider_type=provider_type)
        self.profiles[provider_id] = profile
        return profile

    def active_leader(self, on_date: date) -> LeadershipAssignment | None:
        acting = [x for x in self.leadership if x.role == LeadershipRole.ACTING_MANAGER and x.is_active(on_date)]
        if len(acting) > 1:
            raise ValueError(f"More than one acting manager is active on {on_date.isoformat()}")
        if acting:
            return acting[0]
        managers = [x for x in self.leadership if x.role == LeadershipRole.MANAGER and x.is_active(on_date)]
        if len(managers) > 1:
            raise ValueError(f"More than one manager is active on {on_date.isoformat()}")
        return managers[0] if managers else None

    def effective_state(self, provider_id: str, on_date: date) -> EffectiveProviderState:
        profile = self.ensure_profile(provider_id)
        leader = self.active_leader(on_date)
        leadership_role = leader.role if leader and leader.provider_id == provider_id else None
        active_overrides = [x for x in self.overrides if x.provider_id == provider_id and x.is_active(on_date)]
        blocked_times = frozenset(x.time_slot for x in active_overrides
                                  if x.kind == OverrideKind.BLOCK_TIME and x.time_slot is not None)
        open_times = frozenset(x.time_slot for x in active_overrides
                               if x.kind == OverrideKind.OPEN_REPLACEMENT_TIME and x.time_slot is not None)
        max_slots = profile.normal_max_slots
        max_overrides = [x.max_load for x in active_overrides
                         if x.kind == OverrideKind.MAX_LOAD and x.max_load is not None]
        if max_overrides:
            max_slots = min(max_overrides)
        replacement_eligible = profile.replacement_eligible and not any(
            x.kind == OverrideKind.EXCLUDE_REPLACEMENTS for x in active_overrides
        )
        if leadership_role is not None:
            max_slots = 0
            replacement_eligible = False
        return EffectiveProviderState(
            provider_id=profile.provider_id,
            display_name=profile.display_name,
            active=profile.active,
            provider_type=profile.provider_type,
            leadership_role=leadership_role,
            max_slots=max_slots,
            replacement_eligible=replacement_eligible,
            blocked_times=blocked_times,
            explicitly_open_replacement_times=open_times,
            applied_override_reasons=tuple(x.reason for x in active_overrides if x.reason),
        )

    def can_take_replacement(self, *, provider_id: str, on_date: date,
                             requested_time: time, occupied_times: Iterable[time]) -> PolicyDecision:
        state = self.effective_state(provider_id, on_date)
        occupied = set(occupied_times)
        if not state.active:
            return PolicyDecision(False, "provider_inactive", state)
        if requested_time in state.blocked_times:
            return PolicyDecision(False, "time_blocked", state)
        if requested_time in occupied:
            return PolicyDecision(False, "time_already_occupied", state)
        if state.is_active_leader:
            if requested_time not in state.explicitly_open_replacement_times:
                return PolicyDecision(False, "leader_time_not_explicitly_open", state)
            return PolicyDecision(True, "leader_explicit_time_open", state)
        if not state.replacement_eligible:
            return PolicyDecision(False, "excluded_from_replacements", state)
        if len(occupied) >= state.max_slots:
            return PolicyDecision(False, "capacity_reached", state)
        return PolicyDecision(True, "allowed", state)

    def can_take_permanent_assignment(self, *, provider_id: str, on_date: date,
                                      requested_time: time, occupied_times: Iterable[time]) -> PolicyDecision:
        state = self.effective_state(provider_id, on_date)
        occupied = set(occupied_times)
        if not state.active:
            return PolicyDecision(False, "provider_inactive", state)
        if state.is_active_leader:
            return PolicyDecision(False, "active_leader_has_zero_base_load", state)
        if requested_time in state.blocked_times:
            return PolicyDecision(False, "time_blocked", state)
        if requested_time in occupied:
            return PolicyDecision(False, "time_already_occupied", state)
        if len(occupied) >= state.max_slots:
            return PolicyDecision(False, "capacity_reached", state)
        return PolicyDecision(True, "allowed", state)
