"""Publish verified report outputs without overwriting or deleting user files.

Writers receive private staging paths in the SAME destination filesystem.
The context manager publishes stages only when the writer returns normally.
Publication uses atomic exclusive hard links, so a pre-existing destination
(including one created after preflight) cannot be overwritten. If publishing
one of several report outputs fails, links created by this invocation are
removed; pre-existing or replaced files are never deleted. Staging files are
always cleaned. Multi-file publication is not an indivisible atomic operation;
only independently verified outputs are ever exposed.

NTFS/local POSIX hard-link support is required. Unsupported/network volumes
fail closed rather than falling back to unsafe overwrite/rename semantics.
"""
from __future__ import annotations

from contextlib import ExitStack, contextmanager
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterator


class VerifiedReportPublicationError(ValueError):
    """A verified report could not be safely published."""


@contextmanager
def staged_report_outputs(
    destinations: tuple[str | Path, ...],
) -> Iterator[tuple[Path, ...]]:
    """Yield unused private paths; publish all on normal exit, never clobber.

    Callers must completely write and verify each staged output inside the
    context. Raising any exception prevents publishing. Pre-existing outputs
    are the caller's property, including when a previous run failed.
    """
    if not destinations:
        raise VerifiedReportPublicationError("At least one output is required")
    # Normalize dot segments without following a destination symlink, which
    # could otherwise redirect a report to a different user-owned location.
    targets = tuple(Path(os.path.abspath(path)) for path in destinations)
    if len(set(targets)) != len(targets):
        raise VerifiedReportPublicationError("Report output destinations collide")
    if any(path.exists() or path.is_symlink() for path in targets):
        raise VerifiedReportPublicationError("Report output already exists")

    published: list[tuple[Path, Path]] = []
    with ExitStack() as staging:
        stages: list[Path] = []
        for target in targets:
            target.parent.mkdir(parents=True, exist_ok=True)
            stage_dir = Path(staging.enter_context(
                TemporaryDirectory(prefix=".diomil_unverified_", dir=target.parent)
            ))
            stages.append(stage_dir / ("stage" + target.suffix))

        try:
            yield tuple(stages)
            for stage in stages:
                if not stage.is_file():
                    raise VerifiedReportPublicationError("Verified staged report missing")
            for stage, target in zip(stages, targets):
                # os.link is exclusive: EEXIST is a safety stop, not overwrite.
                os.link(stage, target)
                published.append((stage, target))
        except BaseException:
            for stage, target in reversed(published):
                try:
                    if target.exists() and os.path.samefile(stage, target):
                        target.unlink()
                except OSError:
                    # Avoid deleting an external/replaced user file. A retained
                    # link is an already-verified output, not unverified bytes.
                    pass
            raise
