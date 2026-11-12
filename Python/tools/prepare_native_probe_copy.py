"""Create one exclusive, byte-exact local XLSM probe copy; never overwrite.

The source may contain clinical data. Do not upload either file to GitHub or
external storage. This builder does NOT open native Excel or save a workbook.
"""
from __future__ import annotations

import argparse
from hashlib import sha256
from pathlib import Path
from zipfile import BadZipFile, ZipFile


class ProbeCopyError(RuntimeError):
    """Stop before a protected workbook can be overwritten."""


def _sha256(path: Path) -> str:
    result = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            result.update(chunk)
    return result.hexdigest()


def create_probe_copy(source_path: Path, output_path: Path) -> dict[str, object]:
    """Exclusive-create a single offline probe copy and verify both hashes.

    The output is unconditionally a *new* file. If exclusive opening fails,
    never delete the file that another process created. If a read or write
    fails after this process created the copy, delete only our own partial file.
    """
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output or source.suffix.lower() != ".xlsm" or output.suffix.lower() != ".xlsm":
        raise ProbeCopyError("Distinct source and destination .xlsm paths required")
    if not source.is_file():
        raise ProbeCopyError("Source not found")
    if output.exists():
        raise ProbeCopyError("Destination already exists: refusing overwrite")
    before = _sha256(source)
    try:
        with ZipFile(source) as archive:
            if "xl/vbaProject.bin" not in archive.namelist():
                raise ProbeCopyError("VBA project is missing")
    except BadZipFile as exc:
        raise ProbeCopyError("Source is not a valid XLSM") from exc

    if not output.parent.is_dir():
        raise ProbeCopyError("Destination parent folder must already exist")
    created = False
    try:
        with output.open("xb") as sink:
            created = True
            with source.open("rb") as original:
                for chunk in iter(lambda: original.read(1024 * 1024), b""):
                    sink.write(chunk)
        if _sha256(source) != before or _sha256(output) != before:
            raise ProbeCopyError("Read-back hash mismatch")
        with ZipFile(output) as archive:
            if "xl/vbaProject.bin" not in archive.namelist():
                raise ProbeCopyError("Copy lost VBA project")
    except Exception:
        if created:
            output.unlink(missing_ok=True)
        raise
    return {"byte_exact": True, "source_unchanged": True, "copy_created": True}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Safe exclusive-copy-only XLSM native probe setup")
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)
    try:
        result = create_probe_copy(args.source, args.output)
    except Exception as exc:
        # Never leak confidential file paths in aggregate operator logs.
        print(f"SAFETY STOP: Probe copy refused ({type(exc).__name__}).")
        return 2
    if result["byte_exact"] and result["source_unchanged"]:
        print("ISOLATED XLSM BYTE-EXACT COPY OK")
        print("Source unchanged: YES; no overwrite; no Excel session launched.")
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
