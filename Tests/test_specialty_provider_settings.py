from rehab_excel.specialty_provider_settings import collect_specialty_providers


def test_collect_specialty_providers_deduplicates_and_preserves_order():
    headers = {
        "Εργο_Θεραπευτής": 0,
        "Λογο_Θεραπευτής": 1,
        "Ψυχ_Ψυχολόγος": 2,
    }
    rows = (
        ("Εργο 1", "Λογο 1", "Ψυχ 1"),
        ("Εργο 2", "Λογο 1", "Ψυχ 2"),
        (" εργο 1 ", "", None),
    )

    result = collect_specialty_providers(headers, rows)

    assert result["Εργο_Θεραπευτής"] == ("Εργο 1", "Εργο 2")
    assert result["Λογο_Θεραπευτής"] == ("Λογο 1",)
    assert result["Ψυχ_Ψυχολόγος"] == ("Ψυχ 1", "Ψυχ 2")


def test_collect_specialty_providers_handles_missing_planner_columns():
    result = collect_specialty_providers({"Εργο_Θεραπευτής": 0}, (("Εργο 1",),))

    assert result["Εργο_Θεραπευτής"] == ("Εργο 1",)
    assert result["Λογο_Θεραπευτής"] == ()
    assert result["Ψυχ_Ψυχολόγος"] == ()
