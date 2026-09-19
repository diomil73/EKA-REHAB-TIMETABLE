from __future__ import annotations

from dataclasses import dataclass

from .replacement_options import ReplacementProviderOption


@dataclass(frozen=True)
class ReplacementChoice:
    rank: int
    option: ReplacementProviderOption


def choose_replacement_option(
    options: list[ReplacementProviderOption] | tuple[ReplacementProviderOption, ...],
    rank: int,
) -> ReplacementChoice:
    """Return a 1-based ranked replacement choice.

    Keeping rank selection in the core makes the interactive CLI thin and gives
    future Excel UI code the same deterministic selection rule.
    """

    if rank < 1:
        raise ValueError("Replacement rank must be at least 1")
    if rank > len(options):
        raise ValueError(
            f"Replacement rank {rank} is out of range; available choices: {len(options)}"
        )
    return ReplacementChoice(rank=rank, option=options[rank - 1])
