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
        Protection=protection, VBComponents=Items(*components),
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
    frm = workbook.VBProject.VBComponents.Item(0).Designer
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
    designer = obj.VBProject.VBComponents.Item(0).Designer
    del designer.Controls.Item(0).Width
    with pytest.raises(UserFormDesignerAuditError, match="Required designer property"):
        snapshot(obj)

    obj = sample()
    obj.VBProject.VBComponents.Item(0).Designer.Controls.Count = -1
    with pytest.raises(UserFormDesignerAuditError, match="safe bounds"):
        snapshot(obj)

    obj = sample()
    designer = obj.VBProject.VBComponents.Item(0).Designer
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
