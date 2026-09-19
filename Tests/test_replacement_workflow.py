from datetime import time

import pytest

from rehab_core.models import ReplacementProviderKind
from rehab_core.replacement_options import ReplacementProviderOption
from rehab_core.replacement_workflow import choose_replacement_option


def _option(name: str, hour: int) -> ReplacementProviderOption:
    slot = time(hour, 0)
    return ReplacementProviderOption(
        provider_id=name,
        display_name=name,
        provider_kind=ReplacementProviderKind.THERAPIST,
        active_sessions=1,
        infectious_sessions=0,
        robotic_sessions=0,
        replacement_sessions=0,
        requested_time=time(8, 30),
        recommended_time=slot,
        exact_time_available=False,
        available_timeslots=(slot,),
    )


def test_choose_replacement_option_uses_one_based_rank():
    options = [_option("A", 9), _option("B", 10)]
    choice = choose_replacement_option(options, 2)
    assert choice.rank == 2
    assert choice.option.provider_id == "B"


def test_choose_replacement_option_rejects_zero():
    with pytest.raises(ValueError):
        choose_replacement_option([_option("A", 9)], 0)


def test_choose_replacement_option_rejects_rank_past_end():
    with pytest.raises(ValueError):
        choose_replacement_option([_option("A", 9)], 2)
