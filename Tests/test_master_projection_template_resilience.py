from rehab_excel.master_projection_writer import Win32ComMasterProjectionBackend


class _Rows:
    Count = 1000


class _Cell:
    def __init__(self, ws, row, col):
        self.ws = ws
        self.row = row
        self.col = col

    @property
    def Value(self):
        return self.ws.values.get((self.row, self.col))

    @property
    def Formula(self):
        return self.Value

    def End(self, _direction):
        rows = [
            row
            for (row, col), value in self.ws.values.items()
            if col == self.col and value not in (None, "")
        ]
        return type("EndCell", (), {"Row": max(rows or [1])})()


class _Worksheet:
    def __init__(self):
        self.Rows = _Rows()
        self.values = {
            (4, 2): "Θάλαμος",
            (4, 3): "Ασθενής",
            (4, 11): "ΨΥΧΟΛΟΓΟΙ",
            (4, 12): "ΑΠΟΓΕΥΜΑΤΙΝΟ ΠΡΟΓΡΑΜΜΑ",
            (4, 13): "Κατάσταση",
            # Row 5 is intentionally blank: it is the clinic banner row.
            (6, 1): "=PATIENT_PLANNER!D12",
            (6, 2): "=PATIENT_PLANNER!B12",
            # Column 3 can be rendered text when a responsible doctor exists.
            (6, 3): "TEST PATIENT\n✚ TEST DOCTOR",
            (6, 4): "=PATIENT_PLANNER!F12",
            (6, 5): "=PATIENT_PLANNER!I12",
            (6, 6): "=PATIENT_PLANNER!N12",
            (6, 7): "=PATIENT_PLANNER!L12",
            (6, 8): "=PATIENT_PLANNER!P12",
            (6, 9): "=PATIENT_PLANNER!R12",
            (6, 10): "=PATIENT_PLANNER!T12",
            (6, 13): "=PATIENT_PLANNER!E12",
        }

    def Cells(self, row, col):
        return _Cell(self, row, col)


def test_master_formula_templates_survive_blank_clinic_banner_row():
    ws = _Worksheet()

    template_row, formulas = Win32ComMasterProjectionBackend._template_formulas(ws)

    assert template_row == 5
    assert formulas[1] == "=PATIENT_PLANNER!D12"
    assert formulas[2] == "=PATIENT_PLANNER!B12"
    assert formulas[4] == "=PATIENT_PLANNER!F12"
    assert formulas[13] == "=PATIENT_PLANNER!E12"
    # Patient identity may be display text, so the stable mapping supplies C.
    assert formulas[3] == "=PATIENT_PLANNER!C2"


def test_canonical_master_formula_contract_supports_expanded_status_column():
    formulas = Win32ComMasterProjectionBackend._canonical_formula_templates(13)

    assert formulas[1] == "=PATIENT_PLANNER!D2"
    assert formulas[3] == "=PATIENT_PLANNER!C2"
    assert formulas[10] == "=PATIENT_PLANNER!T2"
    assert formulas[13] == "=PATIENT_PLANNER!E2"
