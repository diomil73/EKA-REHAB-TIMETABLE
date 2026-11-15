from __future__ import annotations

"""Non-identifying, read-only native Excel VBA component inventory.

Opens each XLSM read-only with macros/events/links disabled. Reports component
names/types, code line counts, and UserForm control names/types only. It never
prints VBA source, cell values, formulas, captions, or patient/staff data.
"""

import argparse
from hashlib import sha256
from pathlib import Path
import sys


COMPONENT_TYPES = {
    1: "STANDARD_MODULE",
    2: "CLASS_MODULE",
    3: "USERFORM",
    100: "DOCUMENT_MODULE",
}


def component_type_label(value: int) -> str:
    return COMPONENT_TYPES.get(int(value), f"OTHER_TYPE_{int(value)}")


def file_sha256(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _type_name(control) -> str:
    try:
        return str(control._oleobj_.GetTypeInfo().GetDocumentation(-1)[0])
    except Exception:
        return "TYPE_UNAVAILABLE"


def audit_one(path: Path) -> tuple[int, str | None]:
    path = path.expanduser().resolve()
    if sys.platform != "win32" or not path.is_file() or path.suffix.casefold() != ".xlsm":
        print("SAFETY STOP: existing XLSM on Windows is required")
        return 2, None

    try:
        before_hash = file_sha256(path)
        before_size = path.stat().st_size
    except OSError as exc:
        print(f"SAFETY STOP: cannot fingerprint workbook ({type(exc).__name__})")
        return 2, None

    excel = None
    workbook = None
    complete = True
    status = 0
    stage = "importing Excel COM client"
    try:
        import win32com.client  # type: ignore[import-not-found]

        stage = "starting isolated Excel instance"
        excel = win32com.client.DispatchEx("Excel.Application")
        stage = "configuring Excel security and read-only session"
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        # Calculation is an optional Application-level setting. Some localized
        # Excel/COM installations reject this property; do not let that prevent
        # a strictly read-only inventory. No workbook is saved.
        try:
            stage = "setting optional manual calculation mode"
            excel.Calculation = -4135  # xlCalculationManual
            calculation_manual = True
        except Exception as exc:
            calculation_manual = False
            print(
                "WARNING: manual calculation mode unavailable:",
                type(exc).__name__,
            )
        # msoAutomationSecurityForceDisable: do not execute workbook macros.
        stage = "setting macro security"
        excel.AutomationSecurity = 3

        stage = "opening workbook read-only"
        workbook = excel.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=True,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        stage = "verifying workbook read-only state"
        if not bool(workbook.ReadOnly):
            print("SAFETY STOP: workbook did not open read-only")
            status = 2
        else:
            stage = "accessing VBA project object model"
            project = workbook.VBProject
            stage = "checking VBA project protection"
            protection = int(project.Protection)
            if protection != 0:
                print("SAFETY STOP: VBA project is locked; component inventory unavailable")
                status = 2
            else:
                stage = "enumerating VBA components"
                components = project.VBComponents
                count = int(components.Count)
                print("DIOMIL NATIVE VBA PROJECT INVENTORY")
                print("Candidate:", path.name)
                print("Workbook SHA256:", before_hash)
                print("Workbook size bytes:", before_size)
                print("Excel ReadOnly:", bool(workbook.ReadOnly))
                print("Macros disabled before open: YES")
                print("Excel events disabled: YES")
                print("Calculation mode set to manual:", "YES" if calculation_manual else "UNAVAILABLE")
                print("VBA project protection:", protection)
                print("VBA component count:", count)

                for index in range(1, count + 1):
                    component = components.Item(index)
                    name = str(component.Name)
                    component_type = int(component.Type)
                    label = component_type_label(component_type)
                    code_lines = int(component.CodeModule.CountOfLines)
                    print(
                        f"COMPONENT {index:03d} | name={name} | "
                        f"type={label}({component_type}) | code_lines={code_lines}"
                    )

                    if component_type == 3:
                        stage = f"enumerating controls for UserForm {name}"
                        try:
                            controls = component.Designer.Controls
                            control_count = int(controls.Count)
                            print(f"  USERFORM_CONTROLS: {control_count}")
                            for control_index in range(0, control_count):
                                control = controls.Item(control_index)
                                control_name = str(control.Name)
                                control_type = _type_name(control)
                                if control_type == "TYPE_UNAVAILABLE":
                                    complete = False
                                print(
                                    f"  CONTROL {control_index + 1:03d} | "
                                    f"name={control_name} | type={control_type}"
                                )
                        except Exception as exc:
                            complete = False
                            print(
                                "  USERFORM CONTROL INVENTORY INCOMPLETE:",
                                type(exc).__name__,
                            )

                print("VBA SOURCE TEXT PRINTED: NO")
                print("CELL VALUES OR FORMULAS PRINTED: NO")

    except ImportError:
        print("SAFETY STOP: pywin32 is required for native Excel VBA inventory")
        status = 2
    except Exception as exc:
        print(
            f"SAFETY STOP: native VBA inventory failed at stage={stage} "
            f"({type(exc).__name__})"
        )
        hresult = getattr(exc, "hresult", None)
        if isinstance(hresult, int):
            print(f"COM HRESULT: 0x{hresult & 0xFFFFFFFF:08X}")
        excepinfo = getattr(exc, "excepinfo", None)
        if isinstance(excepinfo, tuple) and len(excepinfo) > 2 and excepinfo[2]:
            print("COM DESCRIPTION:", str(excepinfo[2])[:500])
        status = 2
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                complete = False
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                complete = False

    try:
        after_hash = file_sha256(path)
        after_size = path.stat().st_size
    except OSError as exc:
        print(f"SAFETY STOP: cannot verify source after inspection ({type(exc).__name__})")
        return 2, None

    unchanged = (before_hash, before_size) == (after_hash, after_size)
    print("SOURCE HASH UNCHANGED:", "YES" if unchanged else "NO")
    if not unchanged:
        print("SAFETY STOP: source workbook changed during read-only inspection")
        return 2, after_hash
    if status != 0:
        print("NATIVE VBA PROJECT INVENTORY NOT COMPLETED")
        return status, after_hash
    if not complete:
        print("NATIVE VBA PROJECT INVENTORY INCOMPLETE")
        return 3, after_hash
    print("NATIVE VBA PROJECT INVENTORY COMPLETE")
    print("NO WORKBOOK WAS MODIFIED")
    return 0, after_hash

def audit_many(paths: list[Path]) -> int:
    seen: dict[str, tuple[str, int]] = {}
    results: list[int] = []
    for path in paths:
        print("")
        print("=== VBA CANDIDATE START ===")
        try:
            fingerprint = file_sha256(path.expanduser().resolve())
        except OSError:
            fingerprint = None
        if fingerprint is not None and fingerprint in seen:
            previous_name, previous_code = seen[fingerprint]
            print("Candidate:", path.name)
            print("DUPLICATE BYTE-IDENTICAL CONTENT; already audited:", previous_name)
            print("VBA component inventory skipped for duplicate")
            print("Original audit status:", previous_code)
            results.append(previous_code)
            print("=== VBA CANDIDATE END ===")
            continue

        code, resulting_hash = audit_one(path)
        results.append(code)
        if fingerprint is not None:
            seen[fingerprint] = (path.name, code)
        if resulting_hash is not None and resulting_hash != fingerprint:
            seen[resulting_hash] = (path.name, code)
        print("=== VBA CANDIDATE END ===")

    if all(code == 0 for code in results):
        print("ALL DISTINCT VBA PROJECT INVENTORIES COMPLETE")
        return 0
    print("VBA PROJECT INVENTORIES REQUIRE REVIEW:", results)
    return 3 if 3 in results else 2

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Non-identifying, read-only native Excel VBA component inventory."
    )
    parser.add_argument("workbooks", nargs="+", type=Path)
    args = parser.parse_args(argv)
    return audit_many(args.workbooks)


if __name__ == "__main__":
    raise SystemExit(main())
