from __future__ import annotations

import sys


def main() -> int:
    if sys.platform != "win32":
        print("NOT READY: native Excel write-back requires Windows.")
        return 2
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        print("NOT READY: install pywin32 with: pip install pywin32")
        return 2

    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        print(f"READY: Microsoft Excel native automation available (version {excel.Version}).")
        return 0
    except Exception as exc:
        print(f"NOT READY: could not start Microsoft Excel: {exc}")
        return 2
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
