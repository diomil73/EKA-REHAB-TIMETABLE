"""Synthetic-only contract tests for A:V weekday projection. No real records."""
from datetime import time

import pytest

from rehab_excel.trial_source_weekday_projection import (
    TrialSourceWeekdayError, preview_source_weekdays,
)


def source_row(**cells):
    values = [None] * 22
    values[1] = "DEMO PATIENT"
    for column, value in cells.items():
        values[ord(column) - ord("A")] = value
    return values


def preview(**cells):
    return preview_source_weekdays(
        source_row=4,
        columns=source_row(**cells),
        confirmed_patient_id="DEMO-1",
    )


def test_blank_weekdays_with_time_expand_exactly_monday_to_friday():
    result = preview(C="09:30", E="DEMO THERAPIST")
    assert result.ready_for_independent_review
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell) for x in result.entries] == [
        (day, time(9, 30), "C4") for day in ("ΔΕ", "ΤΡ", "ΤΕ", "ΠΕ", "ΠΑ")
    ]


def test_blank_clock_and_blank_weekdays_generate_no_phantom_sessions():
    result = preview()
    assert result.entries == ()
    assert result.blockers == ()
    assert not result.ready_for_independent_review


def test_explicit_weekdays_are_not_expanded_or_changed():
    result = preview(C="08:15", D="ΤΡ-ΠΕ", E="DEMO THERAPIST")
    assert [x.weekday for x in result.entries] == ["ΤΡ", "ΠΕ"]
    assert all(x.start_time == time(8, 15) for x in result.entries)


def test_multiple_time_groups_retain_separate_weekday_reservations():
    result = preview(
        G="09:00/10:00", H="ΔΕ-ΤΕ/ΤΡ-ΠΕ", I="DEMO OCCUPATIONAL"
    )
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell) for x in result.entries] == [
        ("ΔΕ", time(9), "G4"),
        ("ΤΕ", time(9), "G4"),
        ("ΤΡ", time(10), "G4"),
        ("ΠΕ", time(10), "G4"),
    ]


def test_missing_clock_with_day_or_provider_blocks_instead_of_creating_slot():
    result = preview(D="ΔΕ-ΤΕ", E="DEMO THERAPIST")
    assert result.entries == ()
    assert result.blockers == ("MISSING_CLOCK_C",)


def test_unknown_clock_or_weekend_day_blocks_without_leaking_cell_text():
    invalid_clock = preview(C="UNRECOGNIZED PATIENT NOTE", E="DEMO THERAPIST")
    weekend = preview(C="09:30", D="ΣΑ", E="DEMO THERAPIST")
    assert invalid_clock.blockers == ("UNPARSEABLE_TREATMENT_C",)
    assert weekend.blockers == ("UNPARSEABLE_TREATMENT_C",)
    assert "PATIENT NOTE" not in repr(invalid_clock)
    assert invalid_clock.entries == weekend.entries == ()


def test_one_blocked_stream_blocks_entire_source_row_not_just_that_stream():
    result = preview(C="09:30", E="DEMO THERAPIST", U="PM NOTE")
    assert result.entries == ()
    assert result.blockers == ("AFTERNOON_UNMAPPED",)
    assert "DEMO-1" not in repr(result)
    assert "PM NOTE" not in repr(result)


def test_reclining_and_pool_wait_for_specific_operational_contracts():
    result = preview(F="11:00", V="POOL SOURCE TEXT")
    assert result.entries == ()
    assert result.blockers == (
        "RECLINING_REQUIRES_SEPARATE_30_MIN_CONTRACT",
        "POOL_UNMAPPED",
    )


def test_missing_fth_provider_is_not_assumed():
    result = preview(C="09:30")
    assert result.entries == ()
    assert result.blockers == ("UNPARSEABLE_TREATMENT_C",)


def test_blank_context_is_not_forward_filled():
    result = preview(C="09:30", E="DEMO THERAPIST")
    assert len(result.entries) == 5


@pytest.mark.parametrize("bad_row", [0, 1, True])
def test_invalid_source_coordinates_fail_closed(bad_row):
    with pytest.raises(TrialSourceWeekdayError, match="Invalid source row"):
        preview_source_weekdays(
            source_row=bad_row, columns=source_row(C="09:30"),
            confirmed_patient_id="DEMO-1",
        )


def test_bad_source_shape_or_missing_confirmed_id_fails_closed():
    with pytest.raises(TrialSourceWeekdayError, match="Invalid source row"):
        preview_source_weekdays(
            source_row=4, columns=[None] * 21,
            confirmed_patient_id="DEMO-1",
        )
    with pytest.raises(TrialSourceWeekdayError, match="Invalid source row"):
        preview_source_weekdays(
            source_row=4, columns=source_row(C="09:30"),
            confirmed_patient_id="",
        )

@pytest.mark.parametrize("clock", ["9\u038415", "09\u038415", "09:15", "09.15", "09,15"])
def test_clock_separator_variants_preserve_exact_weekday_and_source(clock):
    result = preview(G=clock, H="ΤΡ-ΠΕ", I="DEMO STAFF")
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell, x.therapist_name)
            for x in result.entries] == [
        (day, time(9, 15), "G4", "DEMO STAFF") for day in ("ΤΡ", "ΠΕ")
    ]


@pytest.mark.parametrize("clock", ["*", "9\u03841", "9\u038460", "24\u038415",
                                  "9\u038415*", "9\u0384\u038415"])
def test_ambiguous_or_invalid_tonos_clock_remains_blocked(clock):
    result = preview(G=clock, H="ΤΡ-ΠΕ", I="DEMO STAFF")
    assert result.entries == ()
    assert result.blockers == ("UNPARSEABLE_TREATMENT_G",)

def test_efa_distinct_same_day_times_preserve_both_slots():
    result = preview(P="08:00/09:00", Q="ΔΕ-ΤΕ/ΤΕ")
    assert result.blockers == ()
    assert [(x.weekday, x.start_time, x.source_cell) for x in result.entries] == [
        ("ΔΕ", time(8), "P4"), ("ΤΕ", time(8), "P4"), ("ΤΕ", time(9), "P4")
    ]


@pytest.mark.parametrize("clock,days", [
    ("08:00/08:00", "ΔΕ/ΔΕ"),
    ("08:00", "ΔΕ/ΔΕ"),
    ("08:00", "ΔΕ-ΔΕ"),
])
def test_efa_duplicate_exact_slots_still_block(clock, days):
    result = preview(P=clock, Q=days)
    assert result.entries == ()
    assert result.blockers == ("UNPARSEABLE_TREATMENT_P",)


def test_other_specialties_keep_existing_repeat_day_guard():
    result = preview(G="08:00/09:00", H="ΔΕ/ΔΕ")
    assert result.entries == ()
    assert result.blockers == ("UNPARSEABLE_TREATMENT_G",)

def test_explicit_untimed_dispositions_keep_other_timed_treatments():
    result = preview_source_weekdays(
        source_row=4, confirmed_patient_id="DEMO-1",
        columns=source_row(C="08:00", D="ΔΕ", E="DEMO STAFF",
                           G="*", N="ΤΡ", O="DEMO PSYCH"),
        acknowledged_untimed_fields=frozenset({"G4", "M4"}),
    )
    assert result.blockers == ()
    assert [(x.treatment, x.start_time) for x in result.entries] == [
        ("ΦΘ", time(8))
    ]


@pytest.mark.parametrize("cells,address", [
    ({"C": "*"}, "C4"),
    ({"G": "08:00", "H": "ΔΕ"}, "G4"),
    ({"M": "08:00", "N": "ΔΕ"}, "M4"),
    ({"G": "*"}, "G5"),
])
def test_untimed_disposition_never_suppresses_timed_or_wrong_source(cells, address):
    result = preview_source_weekdays(
        source_row=4, confirmed_patient_id="DEMO-1", columns=source_row(**cells),
        acknowledged_untimed_fields=frozenset({address}),
    )
    assert result.entries == ()
    assert result.blockers

def test_combined_weekly_slot_requires_exact_operator_confirmation():
    from dataclasses import replace
    from rehab_core.programme_mapping import (
        SourceWeeklyEntry, ProgrammeMappingError, validate_weekly_entry_consistency,
    )
    a = SourceWeeklyEntry("DEMO-1", "ΦΘ", "ΔΕ", time(8), "DEMO STAFF", "C4")
    b = SourceWeeklyEntry("DEMO-1", "Ρομποτικό", "ΔΕ", time(8), "DEMO STAFF", "R4")
    approval = frozenset({frozenset({a, b})})
    with pytest.raises(ProgrammeMappingError):
        validate_weekly_entry_consistency([a, b])
    validate_weekly_entry_consistency([a, b], confirmed_combined_slots=approval)
    for entries in ([a], [a, b, a], [a, replace(b, therapist_name="OTHER")],
                    [a, replace(b, source_cell="R5")],
                    [a, replace(b, treatment="Εργο")]):
        with pytest.raises(ProgrammeMappingError):
            validate_weekly_entry_consistency(entries, confirmed_combined_slots=approval)


def test_even_explicit_confirmation_cannot_override_wrong_specialty_or_provider():
    from dataclasses import replace
    from rehab_core.programme_mapping import (
        SourceWeeklyEntry, ProgrammeMappingError, validate_weekly_entry_consistency,
    )
    a = SourceWeeklyEntry("DEMO-1", "ΦΘ", "ΔΕ", time(8), "DEMO STAFF", "C4")
    b = SourceWeeklyEntry("DEMO-1", "Ρομποτικό", "ΔΕ", time(8), "DEMO STAFF", "R4")
    for other in (replace(b, therapist_name="OTHER"), replace(b, therapist_name=None),
                  replace(b, treatment="Εργο"), replace(b, treatment="ΦΘ")):
        with pytest.raises(ProgrammeMappingError):
            validate_weekly_entry_consistency(
                [a, other], confirmed_combined_slots=frozenset({frozenset({a, other})}))



def _fictional_r6_operator_group(
    *, specialty="ΕΡΓΟΘΕΡΑΠΕΙΑ", original_clock="09:15",
    original_days="ΤΡ-ΠΕ", new_day=False,
):
    """Build only fictional R6/A:V cells, never a patient workbook."""
    from openpyxl import Workbook
    from rehab_excel.trial_admission_gate import R6_DECISION_HEADERS

    source = Workbook()
    source.active.title = "ΠΡΟΓΡΑΜΜΑ"
    src = source.active
    treatment_columns = ("G", "H", "I") if specialty == "ΕΡΓΟΘΕΡΑΠΕΙΑ" else ("J", "K", "L")
    for col, value in zip(treatment_columns, (original_clock, original_days, "DEMO STAFF")):
        src[f"{col}4"] = value

    review = Workbook()
    weekly = review.active
    weekly.title = "ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"
    weekly.append(["Clinic", "Row", "Name", "Treatment", "Day",
                   "Clock", "Required", "Optional", "Refs", "Raw", "Status"])
    ledger = review.create_sheet("ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ")
    ledger.append(R6_DECISION_HEADERS)
    refs = " / ".join(f"{col}4" for col in treatment_columns)
    original_before = [("ΤΡ", original_clock), ("ΠΕ", original_clock)]
    if specialty == "ΛΟΓΟΘΕΡΑΠΕΙΑ":
        original_before = [("ΤΡ", original_clock)]
        finals = [("ΤΕ", "10:00")]
    else:
        finals = [("ΔΕ", "08:30"), ("ΤΕ", "08:30")]
        if new_day:
            finals.append(("ΠΑ", "08:30"))
    for i, (after_day, after_clock) in enumerate(finals):
        weekly.append([
            "Α", 4, "DEMO PERSON", specialty, after_day, after_clock,
            None, "NEW STAFF", refs, "FICTIONAL", "PROPOSED_OPERATOR_CORRECTION",
        ])
        if i < len(original_before):
            old_day, old_clock = original_before[i]
            old_state, decision = "HOLD_CONFLICT", "EXPLICIT OPERATOR CHANGE"
            before_provider = "DEMO STAFF"
        else:
            old_day, old_clock = after_day, after_clock
            old_state, decision = None, "EXPLICIT OPERATOR CHANGE — νέα ημέρα"
            before_provider = "NEW STAFF"
        ledger.append([
            decision, 20 + i, 4, "DEMO PERSON", specialty,
            old_day, old_clock, None, before_provider, refs,
            old_state, after_day, after_clock, "NEW STAFF",
            "PROPOSED_OPERATOR_CORRECTION",
        ])
    return source, review


def test_r6_operator_complete_group_preserves_added_weekday():
    from rehab_excel.trial_demo_source_reconciliation import (
        _source_bound_r6_corrected_groups,
    )
    source, review = _fictional_r6_operator_group(new_day=True)
    try:
        original = tuple(
            source.active.cell(4, col).value for col in (7, 8, 9)
        )
        overrides = _source_bound_r6_corrected_groups(review, source.active)
        assert set(overrides) == {(4, "G")}
        assert [
            (therapy, day, clock, provider, cell)
            for therapy, day, clock, provider, cell in overrides[(4, "G")]
        ] == [
            ("Εργο", day, time(8, 30), "NEW STAFF", "G4")
            for day in ("ΔΕ", "ΤΕ", "ΠΑ")
        ]
        assert tuple(source.active.cell(4, col).value for col in (7, 8, 9)) == original
    finally:
        source.close()
        review.close()


def test_r6_operator_moved_speech_group_keeps_source_provenance():
    from rehab_excel.trial_demo_source_reconciliation import (
        _source_bound_r6_corrected_groups,
    )
    source, review = _fictional_r6_operator_group(
        specialty="ΛΟΓΟΘΕΡΑΠΕΙΑ", original_clock="11΄30",
        original_days="ΤΡ",
    )
    try:
        result = _source_bound_r6_corrected_groups(review, source.active)
        assert result[(4, "J")] == (
            ("Λογο", "ΤΕ", time(10), "NEW STAFF", "J4"),
        )
    finally:
        source.close()
        review.close()


def test_r6_ambiguous_original_day_needs_explicit_complete_operator_days():
    from openpyxl import Workbook
    from rehab_excel.trial_demo_source_reconciliation import (
        _source_bound_r6_corrected_groups,
        TrialDemoSourceReconciliationError,
    )
    from rehab_excel.trial_admission_gate import R6_DECISION_HEADERS

    src = Workbook()
    src.active["G4"] = "11.30"
    src.active["H4"] = "ΔΕΥΤ-ΠΕΠΑΡ"
    src.active["I4"] = "DEMO STAFF"
    rev = Workbook()
    weekly = rev.active
    weekly.title = "ΩΡΕΣ_ΑΝΑ_ΗΜΕΡΑ"
    weekly.append(["Clinic", "Row", "Name", "Treatment", "Day",
                   "Clock", "Required", "Optional", "Refs", "Raw", "Status"])
    for day in ("ΔΕ", "ΤΕ", "ΠΑ"):
        weekly.append([
            "Α", 4, "DEMO PERSON", "ΕΡΓΟΘΕΡΑΠΕΙΑ", day,
            "11:30", None, "DEMO STAFF", "G4 / H4 / I4",
            "FICTIONAL", (
                "PROPOSED_OPERATOR_CORRECTION" if day == "ΤΕ"
                else "PROPOSED_EXACT_SOURCE"
            ),
        ])
    ledger = rev.create_sheet("ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ")
    ledger.append(R6_DECISION_HEADERS)
    ledger.append([
        "P10: Ε/Θ ΔΕ-ΤΕ-ΠΑ", 20, 4, "DEMO PERSON",
        "ΕΡΓΟΘΕΡΑΠΕΙΑ", "ΠΕ", "11:30", None, "DEMO STAFF",
        "G4 / H4 / I4", "HOLD_CONFLICT", "ΤΕ", "11:30",
        "DEMO STAFF", "PROPOSED_OPERATOR_CORRECTION",
    ])
    try:
        result = _source_bound_r6_corrected_groups(rev, src.active)
        assert [row[1] for row in result[(4, "G")]] == ["ΔΕ", "ΤΕ", "ΠΑ"]
        ledger["A2"] = "FICTIONAL CHANGE WITHOUT COMPLETE DAY LIST"
        with pytest.raises(
            TrialDemoSourceReconciliationError, match="complete operator decision"
        ):
            _source_bound_r6_corrected_groups(rev, src.active)
    finally:
        src.close()
        rev.close()


def test_r6_operator_original_clock_or_ledger_missing_blocks_group():
    from rehab_excel.trial_demo_source_reconciliation import (
        _source_bound_r6_corrected_groups,
        TrialDemoSourceReconciliationError,
    )
    source, review = _fictional_r6_operator_group(new_day=True)
    try:
        source.active["G4"] = "09:45"
        with pytest.raises(TrialDemoSourceReconciliationError, match="before-values"):
            _source_bound_r6_corrected_groups(review, source.active)
        source.active["G4"] = "09:15"
        review["ΑΠΟΦΑΣΕΙΣ_ΕΠΙΚΑΛΥΨΕΩΝ"].delete_rows(4)
        with pytest.raises(TrialDemoSourceReconciliationError, match="ledger"):
            _source_bound_r6_corrected_groups(review, source.active)
    finally:
        source.close()
        review.close()
