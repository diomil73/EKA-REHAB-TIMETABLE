"""Fictional-only checks of read-only MSForms designer evidence.

Never instantiate real native Excel here. Actual Windows/COM acceptance remains
separate and a passing designer snapshot does not bypass the raw OLE guard.
"""
from types import SimpleNamespace

import pytest

from rehab_excel.vba_form_designer_audit import (
    UserFormDesignerAuditError, audit_userform_designers,
    assert_userform_designers_preserved,
)


class Items:
    def __init__(self, *items):
        self.data = list(items)
        self.Count = len(self.data)

    def Item(self, index):
        return self.data[index]


class VbaItems(Items):
    """VBComponents is one-based; MSForms Controls and Pages are zero-based."""
    def Item(self, index):
        return self.data[index - 1]


def control(name, *, caption="Synthetic", controls=None, pages=None, **other):
    item = SimpleNamespace(
        Name=name, Caption=caption, Left=10.0, Top=15.0,
        Width=100.0, Height=18.0, Visible=True, Enabled=True,
        TabIndex=1, Tag="FICTITIOUS", Class="Forms.CommandButton.1",
        Font=SimpleNamespace(Name="Calibri", Size=11.0, Bold=False),
        **other,
    )
    if controls is not None:
        item.Controls = Items(*controls)
    if pages is not None:
        item.Pages = Items(*pages)
    return item


def form(name, *, caption="Fictional scheduler", controls=()):
    designer = SimpleNamespace(
        Caption=caption, Width=390.0, Height=260.0, Name=name,
        Controls=Items(*controls),
        Font=SimpleNamespace(Name="Calibri", Size=10.0),
    )
    return SimpleNamespace(Name=name, Type=3, Designer=designer)


def book(*components, protection=0):
    return SimpleNamespace(VBProject=SimpleNamespace(
        Protection=protection, VBComponents=VbaItems(*components),
    ))


def sample():
    return book(
        form(
            "frmFictional",
            controls=(
                control("cmdSave", caption="Save"),
                control("frameAppointments", controls=(
                    control("txtDay", caption="Day"),
                )),
                control("pagesSchedule", pages=(
                    SimpleNamespace(
                        Name="pageMain", Caption="Main", Index=0,
                        Controls=Items(control("btnConfirm", caption="Confirm")),
                    ),
                )),
            ),
        ),
        SimpleNamespace(Name="modFictional", Type=1),
    )


def snapshot(workbook=None):
    return audit_userform_designers(
        sample() if workbook is None else workbook,
        expected_form_names=("frmFictional",),
    )


def test_unchanged_nested_form_design_snapshots_compare_exactly():
    left = snapshot()
    right = snapshot()
    assert left[0].controls == 5  # includes Frame, MultiPage and nested controls
    assert assert_userform_designers_preserved(left, right) == 1
    assert "frmFictional" not in repr(left)


@pytest.mark.parametrize("change", [
    "form-caption", "control-caption", "control-geometry",
    "control-enabled", "control-font", "nested-control-caption",
    "page-caption", "add-control",
])
def test_form_designer_readback_rejects_ui_mutations_without_exposing_data(change):
    old = snapshot()
    workbook = sample()
    frm = workbook.VBProject.VBComponents.Item(1).Designer
    items = frm.Controls.data
    if change == "form-caption":
        frm.Caption = "Private clinical label"
    elif change == "control-caption":
        items[0].Caption = "Private patient identifier"
    elif change == "control-geometry":
        items[0].Left += 1
    elif change == "control-enabled":
        items[0].Enabled = False
    elif change == "control-font":
        items[0].Font.Bold = True
    elif change == "nested-control-caption":
        items[1].Controls.Item(0).Caption = "Private nested label"
    elif change == "page-caption":
        items[2].Pages.Item(0).Caption = "Private page label"
    elif change == "add-control":
        items.append(control("extraControl"))
        frm.Controls.Count += 1
    after = snapshot(workbook)
    with pytest.raises(UserFormDesignerAuditError, match="changed=1") as caught:
        assert_userform_designers_preserved(old, after)
    for secret in ("frmFictional", "Private", "patient", "extraControl"):
        assert secret not in str(caught.value)


def test_missing_or_added_form_is_not_a_compatible_readback():
    old = snapshot()
    added = (
        *old,
        type(old[0])("syntheticother", "0" * 64, 0),
    )
    with pytest.raises(UserFormDesignerAuditError, match="added=1"):
        assert_userform_designers_preserved(old, added)
    with pytest.raises(UserFormDesignerAuditError, match="missing=1"):
        assert_userform_designers_preserved(added, old)


def test_form_source_inventory_mismatch_fails_closed():
    with pytest.raises(UserFormDesignerAuditError, match="inventory changed"):
        snapshot(book(form("frmChanged")))
    with pytest.raises(UserFormDesignerAuditError, match="Protected"):
        snapshot(book(form("frmFictional"), protection=1))


def test_inaccessible_or_malformed_control_and_nested_collection_stops_audit():
    obj = sample()
    designer = obj.VBProject.VBComponents.Item(1).Designer
    del designer.Controls.Item(0).Width
    with pytest.raises(UserFormDesignerAuditError, match="Required control property Width inaccessible"):
        snapshot(obj)

    obj = sample()
    obj.VBProject.VBComponents.Item(1).Designer.Controls.Count = -1
    with pytest.raises(UserFormDesignerAuditError, match="safe bounds"):
        snapshot(obj)

    obj = sample()
    designer = obj.VBProject.VBComponents.Item(1).Designer
    designer.Controls.Item(1).Controls.data.append(
        control("txtDay")  # duplicate within the same Frame
    )
    designer.Controls.Item(1).Controls.Count += 1
    with pytest.raises(UserFormDesignerAuditError, match="Repeated designer control"):
        snapshot(obj)


def test_existing_native_journal_provisioner_checks_design_before_ole_bytes():
    import inspect
    from Python.tools.provision_daily_input_journal_native import (
        provision_daily_input_journal,
    )
    text = inspect.getsource(provision_daily_input_journal)
    assert "original_form_designers = (" in text
    assert "saved_form_designers = audit_userform_designers(" in text
    assert "assert_userform_designers_preserved(" in text
    assert text.index("assert_userform_designers_preserved(") < text.index(
        "assert_form_storage_preserved("
    )
    assert "AutomationSecurity = FORCE_DISABLE_MACROS" in text


class FakeVbePropertyCollection:
    """Real VBIDE fallback: form geometry lives in VBComponent.Properties."""
    def __init__(self, **values):
        self.values = values

    def Item(self, name):
        return SimpleNamespace(Value=self.values[name])


def _form_with_vbe_geometry_fallback():
    workbook = sample()
    component = workbook.VBProject.VBComponents.Item(1)
    component.Properties = FakeVbePropertyCollection(
        Width=component.Designer.Width, Height=component.Designer.Height,
    )
    del component.Designer.Width
    del component.Designer.Height
    return workbook


def test_native_form_width_height_resolve_from_vbcomponent_when_designer_omits_them():
    """A real COM Designer may not publish inherited VBComponent geometry."""
    direct = snapshot()
    fallback = snapshot(_form_with_vbe_geometry_fallback())
    assert assert_userform_designers_preserved(direct, fallback) == 1


def test_native_vbe_geometry_fallback_detects_changed_real_form_size():
    direct = snapshot()
    workbook = _form_with_vbe_geometry_fallback()
    workbook.VBProject.VBComponents.Item(1).Properties.values["Width"] += 12.0
    after = snapshot(workbook)
    with pytest.raises(UserFormDesignerAuditError, match="changed=1"):
        assert_userform_designers_preserved(direct, after)


def test_native_vbe_geometry_fallback_missing_dimension_stops_without_private_names():
    workbook = _form_with_vbe_geometry_fallback()
    workbook.VBProject.VBComponents.Item(1).Properties.values.pop("Height")
    with pytest.raises(
        UserFormDesignerAuditError,
        match=r"Required form property Height unavailable via VBComponent.Properties",
    ) as result:
        snapshot(workbook)
    assert "frmFictional" not in str(result.value)


def test_missing_control_geometry_discloses_only_property_kind_not_control_identity():
    workbook = sample()
    frame = workbook.VBProject.VBComponents.Item(1).Designer.Controls.Item(1)
    del frame.Top
    with pytest.raises(
        UserFormDesignerAuditError, match="Required control property Top inaccessible"
    ) as result:
        snapshot(workbook)
    assert "frameAppointments" not in str(result.value)
    assert "frmFictional" not in str(result.value)


def test_unexpected_com_property_failure_is_not_treated_as_unsupported():
    workbook = sample()
    component = workbook.VBProject.VBComponents.Item(1)

    class UnexpectedAccessFailure:
        @property
        def Name(self):
            return "synthetic"
        @property
        def Left(self):
            raise RuntimeError("fictional private COM diagnostic value")

        Top = 3.0
        Width = 11.0
        Height = 11.0

    component.Designer.Controls.data[0] = UnexpectedAccessFailure()
    with pytest.raises(
        UserFormDesignerAuditError, match="control property Left COM access failed"
    ) as result:
        snapshot(workbook)
    assert "fictional private" not in str(result.value)


@pytest.mark.parametrize("target", ["form", "control", "font"])
def test_unsupported_scalar_diagnostic_identifies_only_static_property(target):
    """Never serialize COM object type, repr, form name, or clinical content."""
    workbook = sample()
    designer = workbook.VBProject.VBComponents.Item(1).Designer

    class PrivateObject:
        def __repr__(self):
            return "SECRET PATIENT CONTENT"

    if target == "form":
        designer.Caption = PrivateObject()
        expected = "form property Caption unsupported scalar type"
    elif target == "control":
        designer.Controls.Item(0).Caption = PrivateObject()
        expected = "control property Caption unsupported scalar type"
    else:
        designer.Font.Size = PrivateObject()
        expected = "font property Size unsupported scalar type"

    with pytest.raises(UserFormDesignerAuditError) as caught:
        snapshot(workbook)
    message = str(caught.value)
    assert message == expected
    assert "SECRET" not in message
    assert "frmFictional" not in message


def test_required_geometry_with_unsupported_scalar_remains_fail_closed():
    workbook = sample()
    workbook.VBProject.VBComponents.Item(1).Designer.Controls.Item(0).Width = object()
    with pytest.raises(
        UserFormDesignerAuditError,
        match="control property Width unsupported scalar type",
    ):
        snapshot(workbook)


@pytest.mark.parametrize("value", ["10.0", "10.000000000000", "10"])
def test_com_decimal_font_size_matches_float_designer_snapshot(value):
    """COM's IFont.Size may arrive as Decimal; same size means same digest."""
    from decimal import Decimal
    baseline = snapshot()
    workbook = sample()
    workbook.VBProject.VBComponents.Item(1).Designer.Font.Size = Decimal(value)
    assert assert_userform_designers_preserved(baseline, snapshot(workbook)) == 1


def test_com_decimal_font_size_change_is_detected():
    from decimal import Decimal
    before = snapshot()
    workbook = sample()
    workbook.VBProject.VBComponents.Item(1).Designer.Font.Size = Decimal("10.5")
    with pytest.raises(UserFormDesignerAuditError, match="changed=1"):
        assert_userform_designers_preserved(before, snapshot(workbook))


@pytest.mark.parametrize("value", ["NaN", "Infinity", "-Infinity", "1e30"])
def test_com_decimal_font_size_nonfinite_or_out_of_bounds_stops(value):
    from decimal import Decimal
    workbook = sample()
    workbook.VBProject.VBComponents.Item(1).Designer.Font.Size = Decimal(value)
    with pytest.raises(UserFormDesignerAuditError, match="Non-finite designer number"):
        snapshot(workbook)


def test_decimal_other_than_font_size_remains_rejected():
    from decimal import Decimal
    workbook = sample()
    workbook.VBProject.VBComponents.Item(1).Designer.Controls.Item(0).Left = Decimal("10")
    with pytest.raises(
        UserFormDesignerAuditError, match="control property Left unsupported scalar type"
    ):
        snapshot(workbook)


def test_rowsource_com_failure_reports_only_numeric_codes_and_stops():
    workbook = sample()
    first = workbook.VBProject.VBComponents.Item(1).Designer.Controls.Item(0)

    class SensitiveComError(Exception):
        hresult = -2147352567  # DISP_E_EXCEPTION
        excepinfo = (0, "SECRET SOURCE", "SECRET PATIENT", None, 0, -2146827284)

    class ReadonlyRowSource:
        @property
        def RowSource(self):
            raise SensitiveComError("PRIVATE PATIENT LABEL")

    first.__class__ = SimpleNamespace  # no-op for already synthetic object
    # Use an isolated, synthetic control with the same design-time attributes.
    synthetic = SimpleNamespace(**vars(first))
    class ControlProxy:
        def __getattr__(self, key):
            if key == "RowSource":
                raise SensitiveComError("PRIVATE PATIENT LABEL")
            return getattr(synthetic, key)
    workbook.VBProject.VBComponents.Item(1).Designer.Controls.data[0] = ControlProxy()
    with pytest.raises(UserFormDesignerAuditError) as result:
        snapshot(workbook)
    message = str(result.value)
    assert message == (
        "control property RowSource COM access failed; "
        "HRESULT=0x80020009; SCODE=0x800A03EC"
    )
    assert "SECRET" not in message
    assert "PRIVATE" not in message
    assert "cmdSave" not in message


def test_rowsource_unknown_com_failure_stays_fail_closed_without_leaking_exception():
    workbook = sample()
    first = workbook.VBProject.VBComponents.Item(1).Designer.Controls.Item(0)
    class Proxy:
        def __getattr__(self, key):
            if key == "RowSource":
                raise RuntimeError("PRIVATE PATIENT CONTENT")
            return getattr(first, key)
    workbook.VBProject.VBComponents.Item(1).Designer.Controls.data[0] = Proxy()
    with pytest.raises(UserFormDesignerAuditError) as result:
        snapshot(workbook)
    assert str(result.value) == "control property RowSource COM access failed"
