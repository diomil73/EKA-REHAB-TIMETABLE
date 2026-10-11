from __future__ import annotations

from collections.abc import Iterable, Mapping


SPECIALTY_PROVIDER_COLUMNS = {
    "Ψυχ_Ψυχολόγος": (9, "PSYCHOLOGISTS"),
    "Εργο_Θεραπευτής": (10, "OCCUPATIONAL_THERAPISTS"),
    "Λογο_Θεραπευτής": (11, "SPEECH_THERAPISTS"),
}


def _clean(value: object) -> str:
    return str(value).strip() if value is not None else ""


def collect_specialty_providers(
    headers: Mapping[str, int],
    rows: Iterable[tuple[object, ...]],
) -> dict[str, tuple[str, ...]]:
    """Collect unique specialty providers from PATIENT_PLANNER row values.

    ``headers`` uses zero-based column indexes. Output keys are planner provider
    headers and preserve first-seen display spelling/order.
    """

    collected: dict[str, list[str]] = {
        header: [] for header in SPECIALTY_PROVIDER_COLUMNS
    }
    seen: dict[str, set[str]] = {header: set() for header in SPECIALTY_PROVIDER_COLUMNS}

    for row in rows:
        for header in SPECIALTY_PROVIDER_COLUMNS:
            column = headers.get(header)
            if column is None or column >= len(row):
                continue
            value = _clean(row[column])
            if not value:
                continue
            key = value.casefold()
            if key in seen[header]:
                continue
            seen[header].add(key)
            collected[header].append(value)

    return {header: tuple(values) for header, values in collected.items()}


def seed_specialty_provider_settings_com(workbook) -> dict[str, tuple[str, ...]]:
    """Seed empty SETTINGS specialty registries from existing planner assignments.

    This is a migration helper for legacy workbooks. Existing non-empty SETTINGS
    registries are never replaced. The caller owns workbook saving and event state.
    """

    settings_ws = workbook.Worksheets("SETTINGS")
    planner_ws = workbook.Worksheets("PATIENT_PLANNER")

    last_col = int(planner_ws.Cells(1, planner_ws.Columns.Count).End(-4159).Column)  # xlToLeft
    headers: dict[str, int] = {}
    for column in range(1, last_col + 1):
        value = _clean(planner_ws.Cells(1, column).Value2)
        if value:
            headers[value] = column

    seeded: dict[str, tuple[str, ...]] = {}
    for planner_header, (settings_column, settings_header) in SPECIALTY_PROVIDER_COLUMNS.items():
        existing_last = int(
            settings_ws.Cells(settings_ws.Rows.Count, settings_column).End(-4162).Row
        )  # xlUp
        existing_values = []
        if existing_last >= 2:
            for row in range(2, existing_last + 1):
                value = _clean(settings_ws.Cells(row, settings_column).Value2)
                if value:
                    existing_values.append(value)
        if existing_values:
            seeded[planner_header] = tuple(existing_values)
            continue

        planner_column = headers.get(planner_header)
        if planner_column is None:
            seeded[planner_header] = ()
            continue

        last_row = int(
            planner_ws.Cells(planner_ws.Rows.Count, planner_column).End(-4162).Row
        )
        values: list[str] = []
        seen: set[str] = set()
        for row in range(2, last_row + 1):
            value = _clean(planner_ws.Cells(row, planner_column).Value2)
            if not value:
                continue
            key = value.casefold()
            if key in seen:
                continue
            seen.add(key)
            values.append(value)

        if _clean(settings_ws.Cells(1, settings_column).Value2) == "":
            settings_ws.Cells(1, settings_column).Value = settings_header
        for offset, value in enumerate(values, start=2):
            settings_ws.Cells(offset, settings_column).Value = value
        seeded[planner_header] = tuple(values)

    return seeded
