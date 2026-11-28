"""Compare VBA components between a pinned frozen source and a disposable psychology preview.

Read-only native Excel COM; never runs macros, saves, or modifies either workbook.
Only structural counts / digests and a verdict are printed.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "Python"))

from rehab_excel.vba_semantic_audit import (
    VbaSemanticAuditError,
    audit_vba_components,
    assert_vba_semantics_preserved,
    FORM_TYPE,
)

from rehab_excel.vba_form_storage import (
    form_storage_evidence, compare_form_storage_evidence, stream_change_categories,
    VbaFormStorageError,
)
from rehab_excel.vba_form_designer_audit import (
    audit_userform_designers, assert_userform_designers_preserved,
    UserFormDesignerAuditError,
)

SOURCE_REL = "Excel/previews/UNIFIED_H10_CC_NATIVE_READONLY_PROBE_20261112_A.xlsm"
PREVIEW_REL = "Excel/previews/H10_CC_PSYCH_SCHEMA_DISPOSABLE_20261128.xlsm"
SOURCE_SHA256 = "1bd5e58e905f91acb55e5a706cd8b4563b5a98c170a9ec60fe76c95617420b13"


def digest(path: Path) -> str:
    hasher = sha256()
    with path.open("rb") as stream:
        for part in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(part)
    return hasher.hexdigest()


def inspect(excel, path: Path):
    workbook = None
    try:
        workbook = excel.Workbooks.Open(
            str(path), UpdateLinks=0, ReadOnly=True,
            IgnoreReadOnlyRecommended=True, AddToMru=False,
        )
        if not bool(workbook.ReadOnly):
            raise RuntimeError("Workbook did not open read-only")
        components = audit_vba_components(workbook)
        forms = tuple(e.name for e in components if e.component_type == FORM_TYPE)
        designers = audit_userform_designers(workbook, expected_form_names=forms)
        return components, designers
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    root = args.root.resolve()
    source = root / SOURCE_REL
    preview = root / PREVIEW_REL
    if sys.platform != "win32":
        print("SAFETY_STOP=WINDOWS_NATIVE_EXCEL_REQUIRED")
        return 2
    if not source.is_file() or not preview.is_file():
        print("SAFETY_STOP=SOURCE_OR_PREVIEW_MISSING")
        return 2
    original_digest = digest(source)
    preview_digest = digest(preview)
    if original_digest != SOURCE_SHA256:
        print("SAFETY_STOP=FROZEN_SOURCE_HASH_MISMATCH")
        return 2
    try:
        import win32com.client  # type: ignore[import-not-found]
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.AutomationSecurity = 3
            excel.EnableEvents = False
            excel.DisplayAlerts = False
            excel.Visible = False
            source_components, source_designers = inspect(excel, source)
            preview_components, preview_designers = inspect(excel, preview)
        finally:
            if excel is not None:
                excel.Quit()
        forms = tuple(e.name for e in source_components if e.component_type == FORM_TYPE)
        if forms != tuple(e.name for e in preview_components if e.component_type == FORM_TYPE):
            raise VbaSemanticAuditError("UserForm identity inventory differs")
        source_ole = form_storage_evidence(source, forms)
        preview_ole = form_storage_evidence(preview, forms)
        changes = compare_form_storage_evidence(source_ole, preview_ole)
        categories = stream_change_categories(source_ole, preview_ole)
        print(f"USERFORMS_AUDITED={len(forms)}")
        for name, count_value in sorted(changes.items()):
            print(f"OLE_{name.upper()}={count_value}")
        for name, count_value in sorted(categories.items()):
            print(f"OLE_{name.upper()}={count_value}")
        designer_count = assert_userform_designers_preserved(source_designers, preview_designers)
        print(f"USERFORM_DESIGNERS_MATCHED={designer_count}")
        count = assert_vba_semantics_preserved(
            source_components,
            preview_components,
            journal_codename=None,
            expect_new_document=False,
        )
        print("VBA_COMPONENT_SEMANTICS=GREEN")
        print(f"COMPONENTS_MATCHED={count}")
        verdict = 0
    except Exception as exc:
        print("VBA_COMPONENT_SEMANTICS=SAFETY_STOP")
        print(f"REASON={type(exc).__name__}: {exc}")
        verdict = 2

    source_unchanged = digest(source) == original_digest
    preview_unchanged = digest(preview) == preview_digest
    print(f"FROZEN_SOURCE_UNCHANGED={source_unchanged}")
    print(f"DISPOSABLE_PREVIEW_UNCHANGED={preview_unchanged}")
    print("NO_MACROS_NO_SAVE=YES")
    return verdict if source_unchanged and preview_unchanged else 2


if __name__ == "__main__":
    raise SystemExit(main())
