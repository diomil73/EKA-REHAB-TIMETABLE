from __future__ import annotations

from dataclasses import dataclass
from datetime import time
from typing import Iterable, Mapping

from .models import ReplacementProviderKind


@dataclass(frozen=True)
class GlobalReplacementCandidate:
    """One already-feasible placement offered to the global assignment solver.

    Hard feasibility stays upstream in the existing replacement candidate/policy
    pipeline. This layer chooses among feasible placements for all displaced
    sessions together so a locally attractive choice cannot block a complete
    daily solution.
    """

    request_id: str
    provider_id: str
    replacement_time: time
    provider_active_timeslots: int = 0
    is_original_time: bool = False
    adjacency_penalty: int = 0
    provider_kind: ReplacementProviderKind = ReplacementProviderKind.THERAPIST

    def __post_init__(self) -> None:
        if not self.request_id:
            raise ValueError("request_id is required")
        if not self.provider_id:
            raise ValueError("provider_id is required")
        if self.provider_active_timeslots < 0:
            raise ValueError("provider_active_timeslots must be >= 0")
        if self.adjacency_penalty < 0:
            raise ValueError("adjacency_penalty must be >= 0")


@dataclass(frozen=True)
class GlobalReplacementRequest:
    """One displaced session that needs at most one replacement placement."""

    request_id: str
    patient_id: str
    candidates: tuple[GlobalReplacementCandidate, ...]

    def __post_init__(self) -> None:
        if not self.request_id:
            raise ValueError("request_id is required")
        if not self.patient_id:
            raise ValueError("patient_id is required")
        if any(candidate.request_id != self.request_id for candidate in self.candidates):
            raise ValueError("candidate request_id must match its request")


@dataclass(frozen=True)
class GlobalReplacementPlan:
    assignments: tuple[GlobalReplacementCandidate, ...]
    unresolved_request_ids: tuple[str, ...]


@dataclass(frozen=True)
class _Score:
    placed: int
    adjacency_penalty: int
    original_time_count: int
    workload_spread: int
    total_load: int
    uses_1300: int
    signature: tuple[tuple[str, str, str], ...]

    def ranking_key(self) -> tuple[object, ...]:
        # Higher is better. The priorities intentionally mirror the agreed
        # business contract: place everyone first; protect treatment adjacency;
        # then preserve original time; then balance workload. 13:00 is weak.
        return (
            self.placed,
            -self.adjacency_penalty,
            self.original_time_count,
            -self.workload_spread,
            -self.total_load,
            -self.uses_1300,
            tuple(tuple(-ord(ch) for ch in "|".join(item)) for item in self.signature),
        )


def _provider_baselines(
    requests: tuple[GlobalReplacementRequest, ...],
) -> dict[str, int]:
    values: dict[str, int] = {}
    for request in requests:
        for candidate in request.candidates:
            previous = values.get(candidate.provider_id)
            if previous is None:
                values[candidate.provider_id] = candidate.provider_active_timeslots
            elif previous != candidate.provider_active_timeslots:
                raise ValueError(
                    "provider_active_timeslots must be consistent for each provider"
                )
    return values


def _score(
    assignments: tuple[GlobalReplacementCandidate, ...],
    provider_baselines: Mapping[str, int],
) -> _Score:
    assigned_counts = {provider_id: 0 for provider_id in provider_baselines}
    for assignment in assignments:
        assigned_counts[assignment.provider_id] = (
            assigned_counts.get(assignment.provider_id, 0) + 1
        )

    final_loads = [
        provider_baselines[provider_id] + assigned_counts.get(provider_id, 0)
        for provider_id in sorted(provider_baselines)
    ]
    spread = max(final_loads) - min(final_loads) if final_loads else 0

    signature = tuple(
        sorted(
            (
                assignment.request_id,
                assignment.provider_id,
                assignment.replacement_time.strftime("%H:%M"),
            )
            for assignment in assignments
        )
    )
    return _Score(
        placed=len(assignments),
        adjacency_penalty=sum(item.adjacency_penalty for item in assignments),
        original_time_count=sum(item.is_original_time for item in assignments),
        workload_spread=spread,
        total_load=sum(final_loads),
        uses_1300=sum(item.replacement_time == time(13, 0) for item in assignments),
        signature=signature,
    )


def _optimistic_min_workload_spread(
    provider_baselines: Mapping[str, int],
    provider_counts: Mapping[str, int],
    remaining_assignments: int,
) -> int:
    """Lower-bound final workload spread by allowing ideal free redistribution.

    This intentionally ignores provider eligibility and slot conflicts, so it can
    only be more optimistic than reality and is therefore safe for pruning.
    """

    if not provider_baselines:
        return 0
    loads = [
        provider_baselines[provider_id] + provider_counts.get(provider_id, 0)
        for provider_id in sorted(provider_baselines)
    ]
    for _ in range(remaining_assignments):
        lowest_index = min(range(len(loads)), key=loads.__getitem__)
        loads[lowest_index] += 1
    return max(loads) - min(loads)


def plan_global_replacements(
    requests: Iterable[GlobalReplacementRequest],
    *,
    provider_capacity_remaining: Mapping[str, int] | None = None,
) -> GlobalReplacementPlan:
    """Choose one global replacement assignment for a concrete service date.

    Inputs are pre-filtered feasible candidates. The solver itself enforces the
    cross-request constraints that local ranking cannot see:

    - one provider cannot receive two displaced patients in the same clock slot,
    - one patient cannot receive two replacement sessions at the same clock slot,
    - provider remaining daily capacity is respected across the whole plan,
    - unresolved is allowed only when no better complete/partial assignment exists.

    Search remains deterministic and exact, but uses safe optimistic bounds to
    avoid enumerating branches that cannot beat the best plan already found.
    """

    request_tuple = tuple(requests)
    if len({request.request_id for request in request_tuple}) != len(request_tuple):
        raise ValueError("request_id values must be unique")

    capacities = dict(provider_capacity_remaining or {})
    if any(value < 0 for value in capacities.values()):
        raise ValueError("provider capacity remaining must be >= 0")

    baselines = _provider_baselines(request_tuple)
    ordered_requests = tuple(
        sorted(
            request_tuple,
            key=lambda request: (len(request.candidates), request.request_id),
        )
    )

    suffix_min_adjacency = [0] * (len(ordered_requests) + 1)
    suffix_original_possible = [0] * (len(ordered_requests) + 1)
    for index in range(len(ordered_requests) - 1, -1, -1):
        candidates = ordered_requests[index].candidates
        min_adjacency = min(
            (candidate.adjacency_penalty for candidate in candidates),
            default=0,
        )
        suffix_min_adjacency[index] = suffix_min_adjacency[index + 1] + min_adjacency
        suffix_original_possible[index] = (
            suffix_original_possible[index + 1]
            + int(any(candidate.is_original_time for candidate in candidates))
        )

    best_assignments: tuple[GlobalReplacementCandidate, ...] = ()
    best_score: _Score | None = None

    def visit(
        index: int,
        selected: tuple[GlobalReplacementCandidate, ...],
        occupied_provider_slots: frozenset[tuple[str, time]],
        occupied_patient_slots: frozenset[tuple[str, time]],
        provider_counts: Mapping[str, int],
        adjacency_so_far: int,
        original_count_so_far: int,
    ) -> None:
        nonlocal best_assignments, best_score

        remaining = len(ordered_requests) - index
        placed_upper_bound = len(selected) + remaining
        if best_score is not None and placed_upper_bound < best_score.placed:
            return

        if best_score is not None and placed_upper_bound == best_score.placed:
            optimistic_adjacency = adjacency_so_far + suffix_min_adjacency[index]
            if optimistic_adjacency > best_score.adjacency_penalty:
                return
            if optimistic_adjacency == best_score.adjacency_penalty:
                optimistic_original = (
                    original_count_so_far + suffix_original_possible[index]
                )
                if optimistic_original < best_score.original_time_count:
                    return
                if optimistic_original == best_score.original_time_count:
                    optimistic_spread = _optimistic_min_workload_spread(
                        baselines,
                        provider_counts,
                        remaining,
                    )
                    if optimistic_spread > best_score.workload_spread:
                        return

        if index == len(ordered_requests):
            score = _score(selected, baselines)
            if best_score is None or score.ranking_key() > best_score.ranking_key():
                best_score = score
                best_assignments = selected
            return

        request = ordered_requests[index]
        candidates = tuple(
            sorted(
                request.candidates,
                key=lambda candidate: (
                    candidate.adjacency_penalty,
                    0 if candidate.is_original_time else 1,
                    candidate.provider_active_timeslots,
                    1 if candidate.replacement_time == time(13, 0) else 0,
                    candidate.replacement_time,
                    candidate.provider_id.casefold(),
                    candidate.provider_id,
                ),
            )
        )

        for candidate in candidates:
            provider_slot = (candidate.provider_id, candidate.replacement_time)
            patient_slot = (request.patient_id, candidate.replacement_time)
            if provider_slot in occupied_provider_slots:
                continue
            if patient_slot in occupied_patient_slots:
                continue

            used = provider_counts.get(candidate.provider_id, 0)
            remaining_capacity = capacities.get(candidate.provider_id)
            if remaining_capacity is not None and used >= remaining_capacity:
                continue

            next_counts = dict(provider_counts)
            next_counts[candidate.provider_id] = used + 1
            visit(
                index + 1,
                selected + (candidate,),
                occupied_provider_slots | {provider_slot},
                occupied_patient_slots | {patient_slot},
                next_counts,
                adjacency_so_far + candidate.adjacency_penalty,
                original_count_so_far + int(candidate.is_original_time),
            )

        # Once a complete plan has been found, this unresolved branch can never
        # tie its placed count and therefore cannot improve the result.
        if best_score is None or best_score.placed < len(ordered_requests):
            visit(
                index + 1,
                selected,
                occupied_provider_slots,
                occupied_patient_slots,
                provider_counts,
                adjacency_so_far,
                original_count_so_far,
            )

    visit(0, (), frozenset(), frozenset(), {}, 0, 0)

    assignments = tuple(sorted(best_assignments, key=lambda item: item.request_id))
    resolved = {assignment.request_id for assignment in assignments}
    unresolved = tuple(
        request.request_id
        for request in sorted(request_tuple, key=lambda item: item.request_id)
        if request.request_id not in resolved
    )
    return GlobalReplacementPlan(
        assignments=assignments,
        unresolved_request_ids=unresolved,
    )
