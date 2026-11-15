"""H10: static workbook-event body hashes must match native read-only evidence.

Hashes originate from DIOMIL_VBA_PROFILE_BUNDLE_20261115_195014.txt.
This test does not access workbook bytes or execute Excel/VBA.
"""

from hashlib import sha256

from rehab_excel.application_shell_vba import _EVENT_SPECS


EXPECTED_EVENT_CALLS = {
    "Workbook_Open": (
        "EnterApplicationShell",
        "587b673406a765abb8912aa2545fdba5bf4a08f3380f613a846cbd064d7c623f",
    ),
    "Workbook_Activate": (
        "EnterApplicationShell",
        "587b673406a765abb8912aa2545fdba5bf4a08f3380f613a846cbd064d7c623f",
    ),
    "Workbook_Deactivate": (
        "ExitApplicationShell",
        "c755c205e4b41217bc504583e6eea3429fb87f9618ce474a1ac3f33694d83b62",
    ),
    "Workbook_SheetActivate": (
        "KeepApplicationShell",
        "04ddae21bbda7b48a8a6556d4035ff716ab8696d1d17d08a8be267399ef3bbcb",
    ),
    "Workbook_BeforeClose": (
        "ExitApplicationShell",
        "c755c205e4b41217bc504583e6eea3429fb87f9618ce474a1ac3f33694d83b62",
    ),
}


def test_native_workbook_event_bodies_match_tracked_shell_generator():
    actual = {name: call for name, _signature, call in _EVENT_SPECS}
    assert set(actual) == set(EXPECTED_EVENT_CALLS)
    for name, (call, observed_hash) in EXPECTED_EVENT_CALLS.items():
        assert actual[name] == call
        # The installer inserts an indented one-line VBA event call.
        generated_body = "    " + actual[name]
        assert sha256(generated_body.encode("utf-8")).hexdigest() == observed_hash
