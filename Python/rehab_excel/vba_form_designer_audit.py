"""Read-only, scoped UserForm designer evidence across native Excel saves.

The historic raw UserForm OLE f/o streams change even between two no-edit
Excel Saves. Inspect the *visible MSForms designer contract* instead, without
instantiating forms, running macros, reading runtime patient data, or saving.

This is a supplementary fail-closed CHECK, not proof of every binary resource,
event behavior, or clinical functionality. It NEVER waives existing OLE checks.
Errors and results report counts only; private form/control values stay in RAM.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from hashlib import sha256
import json

from rehab_excel.vba_semantic_audit import FORM_TYPE


class UserFormDesignerAuditError(RuntimeError):
    """UserForm designer structure or property readback was unsafe."""


@dataclass(frozen=True, repr=False)
class UserFormDesignerEvidence:
    name: str
    digest: str
    controls: int


# Only stable design-time properties. Do NOT inspect Text/Value/ListIndex,
# active selections, Picture/OLE raw resources, or any runtime patient facts.
_FORM_PROPERTIES = (
    "Name", "Caption", "Width", "Height", "BackColor", "ForeColor",
    "BorderStyle", "ScrollBars", "ScrollWidth", "ScrollHeight",
    "StartUpPosition", "ShowModal", "Cycle",
)
_CONTROL_PROPERTIES = (
    "Name", "ProgID", "Class", "Caption", "Tag", "Left", "Top",
    "Width", "Height", "Visible", "Enabled", "TabIndex", "TabStop",
    "ControlTipText", "BackColor", "ForeColor", "BorderStyle",
    "SpecialEffect", "TextAlign", "WordWrap", "AutoSize",
    "MultiLine", "Locked", "MaxLength", "Style", "ListStyle",
    "MatchEntry", "BoundColumn", "ColumnCount", "ColumnWidths",
    "RowSource", "ControlSource", "EnterKeyBehavior",
    "ExitKeyBehavior", "ShowDropButtonWhen", "PictureSizeMode",
    "PictureAlignment", "PicturePosition", "GroupName",
)
_PAGE_PROPERTIES = (
    "Name", "Caption", "Index", "Visible", "Enabled",
    "Left", "Top", "Width", "Height", "BackColor", "ForeColor",
)
_FONT_PROPERTIES = ("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough")
_MAX_ITEMS = 1024
_MAX_DEPTH = 12


def _scalar(value):
    """Canonicalize only literal scalar properties; reject unknown objects."""
    if value is None or isinstance(value, (str, int, bool)):
        return value
    if isinstance(value, float):
        if not (-1e12 < value < 1e12):
            raise UserFormDesignerAuditError("Non-finite designer number")
        return format(value, ".12g")
    raise UserFormDesignerAuditError("Unsupported designer property type")


def _property(owner, name, *, mandatory=False, context="designer"):
    try:
        value = getattr(owner, name)
        # Some Windows COM/pywin32 IFont.Size values are Decimal rather than
        # float. Only this explicit font field accepts that representation.
        # Keep the same bounded, finite 12-significant-digit scalar digest.
        if context == "font" and name == "Size" and type(value) is Decimal:
            if not value.is_finite() or not (Decimal("-1e12") < value < Decimal("1e12")):
                raise UserFormDesignerAuditError("Non-finite designer number")
            return ("VALUE", format(float(value), ".12g"))
        return ("VALUE", _scalar(value))
    except AttributeError:
        if mandatory:
            raise UserFormDesignerAuditError(
                f"Required {context} property {name} inaccessible"
            ) from None
        return ("UNSUPPORTED",)
    except UserFormDesignerAuditError as exc:
        # Keep strict scalar validation, but identify the static, allowlisted
        # property that failed. Never include the underlying COM value/type,
        # which may contain private form content.
        if str(exc) == "Unsupported designer property type":
            raise UserFormDesignerAuditError(
                f"{context} property {name} unsupported scalar type"
            ) from None
        raise
    except Exception as exc:
        # DISP_E_MEMBERNOTFOUND is normal for controls of another MSForms type.
        # All other COM errors, including project access errors, fail closed.
        if not mandatory and getattr(exc, "hresult", None) == -2147352573:
            return ("UNSUPPORTED",)
        # RowSource may raise a COM exception rather than mean "unsupported".
        # Capture numeric HRESULT/EXCEPINFO scode only; never the exception
        # description, source, argument values or private control identities.
        code_suffix = ""
        if context == "control" and name == "RowSource":
            hresult = getattr(exc, "hresult", None)
            if type(hresult) is int:
                code_suffix = f"; HRESULT=0x{hresult & 0xffffffff:08X}"
            details = getattr(exc, "excepinfo", None)
            if type(details) in (tuple, list) and len(details) >= 6:
                scode = details[5]
                if type(scode) is int:
                    code_suffix += f"; SCODE=0x{scode & 0xffffffff:08X}"
        raise UserFormDesignerAuditError(
            f"{context} property {name} COM access failed{code_suffix}"
        ) from None


def _collection(owner, attribute):
    try:
        return getattr(owner, attribute)
    except AttributeError:
        return None
    except Exception as exc:
        if getattr(exc, "hresult", None) == -2147352573:
            return None
        raise UserFormDesignerAuditError("Designer child collection unreadable") from None


def _items(collection):
    if collection is None:
        return ()
    try:
        count = int(collection.Count)
        if count < 0 or count > _MAX_ITEMS:
            raise UserFormDesignerAuditError("Designer collection exceeds safe bounds")
        result = []
        for index in range(count):
            try:
                item = collection.Item(index)
            except AttributeError:
                item = collection(index)
            result.append(item)
        return tuple(result)
    except UserFormDesignerAuditError:
        raise
    except Exception:
        raise UserFormDesignerAuditError("Designer collection enumeration failed") from None


def _properties(obj, names, *, required=(), context="designer"):
    data = {}
    for name in names:
        data[name] = _property(obj, name, mandatory=name in required, context=context)
    font = _collection(obj, "Font")
    if font is not None:
        data["Font"] = {name: _property(font, name, context="font") for name in _FONT_PROPERTIES}
    else:
        data["Font"] = "UNSUPPORTED"
    return data



def _component_form_property(component, name):
    """Read VBE's documented VBComponent.Properties fallback, not a default.

    A missing required designer dimension must never silently become a zero or
    an ignored field. A successful fallback returns the same canonical value
    shape as the direct MSForms Designer property. Errors redact real values.
    """
    if name not in ("Width", "Height"):
        raise UserFormDesignerAuditError("Invalid required form geometry")
    try:
        properties = component.Properties
        try:
            item = properties.Item(name)
        except AttributeError:
            item = properties(name)
        return ("VALUE", _scalar(item.Value))
    except UserFormDesignerAuditError:
        raise
    except Exception:
        raise UserFormDesignerAuditError(
            f"Required form property {name} unavailable via VBComponent.Properties"
        ) from None

def _designer_tree(obj, *, path, kind, depth, visited_paths, rows):
    if depth > _MAX_DEPTH:
        raise UserFormDesignerAuditError("Nested designer exceeds safe depth")
    name_result = _property(obj, "Name", mandatory=True)
    name = name_result[1]
    if not isinstance(name, str) or not name.strip():
        raise UserFormDesignerAuditError("Unnamed designer control")
    key = (*path, kind, name.casefold())
    if key in visited_paths:
        raise UserFormDesignerAuditError("Repeated designer control path")
    visited_paths.add(key)
    required = ("Name",)
    if kind == "control":
        required = ("Name", "Left", "Top", "Width", "Height")
    properties = _properties(
        obj, _CONTROL_PROPERTIES if kind == "control" else _PAGE_PROPERTIES,
        required=required, context=kind,
    )
    rows.append((key, properties))
    # A Frame can contain Controls, a MultiPage contains Pages, and each
    # Page can itself contain Controls. Iterate all, never just top-level UI.
    for child in _items(_collection(obj, "Controls")):
        _designer_tree(
            child, path=key, kind="control", depth=depth+1,
            visited_paths=visited_paths, rows=rows,
        )
    for child in _items(_collection(obj, "Pages")):
        _designer_tree(
            child, path=key, kind="page", depth=depth+1,
            visited_paths=visited_paths, rows=rows,
        )


def audit_userform_designers(workbook, *, expected_form_names):
    """Snapshot designer structure/allowlisted static properties, read-only.

    The authoritative VBA component audit must already have produced the
    exact expected UserForm names. No names/captions are logged or returned by
    a mismatch exception. Images/OLE binary payloads require other proof.
    """
    if (
        not isinstance(expected_form_names, tuple) or not expected_form_names
        or any(not isinstance(x, str) or not x for x in expected_form_names)
        or len({x.casefold() for x in expected_form_names}) != len(expected_form_names)
    ):
        raise UserFormDesignerAuditError("Invalid expected UserForm inventory")
    try:
        project = workbook.VBProject
        if int(project.Protection) != 0:
            raise UserFormDesignerAuditError("Protected VBA project")
        components = project.VBComponents
        count = int(components.Count)
        if count < 1 or count > _MAX_ITEMS:
            raise UserFormDesignerAuditError("Invalid VBA component count")
        forms = {}
        for index in range(1, count + 1):
            component = components.Item(index)
            if int(component.Type) != FORM_TYPE:
                continue
            name = str(component.Name)
            key = name.casefold()
            if not name or key in forms:
                raise UserFormDesignerAuditError("Duplicate or unnamed UserForm")
            forms[key] = component
        if set(forms) != {x.casefold() for x in expected_form_names}:
            raise UserFormDesignerAuditError("UserForm inventory changed")
        evidence = []
        for form_name in expected_form_names:
            component = forms[form_name.casefold()]
            designer = component.Designer
            # VBIDE MSForm components expose some form dimensions through
            # VBComponent.Properties even when their Designer COM dispatch
            # omits those members. Probe the canonical VBIDE property ONLY if
            # designer access advertises no such member. Never guess geometry.
            properties = _properties(designer, _FORM_PROPERTIES, context="form")
            for dimension in ("Width", "Height"):
                if properties[dimension] == ("UNSUPPORTED",):
                    properties[dimension] = _component_form_property(
                        component, dimension,
                    )
            rows = []
            paths = set()
            for control in _items(_collection(designer, "Controls")):
                _designer_tree(
                    control, path=(form_name.casefold(),),
                    kind="control", depth=1,
                    visited_paths=paths, rows=rows,
                )
            payload = json.dumps(
                [properties, rows], sort_keys=True,
                ensure_ascii=True, separators=(",", ":"),
            ).encode("utf-8")
            evidence.append(UserFormDesignerEvidence(
                form_name.casefold(), sha256(payload).hexdigest(),
                sum(row[0][-2] == "control" for row in rows),
            ))
        return tuple(sorted(evidence, key=lambda item: item.name))
    except UserFormDesignerAuditError:
        raise
    except Exception:
        raise UserFormDesignerAuditError("Unable to inspect UserForm designers") from None


def assert_userform_designers_preserved(before, after) -> int:
    """Check all baseline forms and their nested controls with no exceptions."""
    original = {item.name: item for item in before}
    current = {item.name: item for item in after}
    if (
        not original or len(original) != len(before)
        or len(current) != len(after)
    ):
        raise UserFormDesignerAuditError("Invalid UserForm designer evidence")
    changes = (
        len(set(original) - set(current)),
        len(set(current) - set(original)),
        sum(
            original[name].digest != current[name].digest
            or original[name].controls != current[name].controls
            for name in set(original) & set(current)
        ),
    )
    if any(changes):
        raise UserFormDesignerAuditError(
            "UserForm designer readback differs: "
            f"missing={changes[0]} added={changes[1]} changed={changes[2]}"
        )
    return len(before)
