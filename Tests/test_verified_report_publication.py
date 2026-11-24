"""Synthetic fault-injection tests for exclusive staged report publication."""
import os
from pathlib import Path

import pytest

from rehab_excel import verified_report_publication as publication


def _stages(path):
    return tuple(path.glob(".diomil_unverified_*"))


def test_success_publishes_only_complete_staged_tsv(tmp_path):
    target = tmp_path / "fictional_report.tsv"
    with publication.staged_report_outputs((target,)) as (stage,):
        assert not target.exists()
        assert stage.parent != target.parent
        stage.write_bytes(b"FICTIONAL\tCOMPLETE\n")
        assert not target.exists()
    assert target.read_bytes() == b"FICTIONAL\tCOMPLETE\n"
    assert not _stages(tmp_path)


def test_exception_midwrite_never_publishes_partial_data(tmp_path):
    target = tmp_path / "fictional_report.tsv"
    with pytest.raises(RuntimeError, match="INTERRUPTION"):
        with publication.staged_report_outputs((target,)) as (stage,):
            stage.write_bytes(b"PARTIAL PRIVATE DATA")
            raise RuntimeError("INTERRUPTION")
    assert not target.exists()
    assert not _stages(tmp_path)


def test_preexisting_output_is_never_overwritten_or_deleted(tmp_path):
    target = tmp_path / "user-owned.tsv"
    target.write_bytes(b"EXISTING USER FILE")
    with pytest.raises(publication.VerifiedReportPublicationError, match="already exists"):
        with publication.staged_report_outputs((target,)):
            raise AssertionError("Must not yield")
    assert target.read_bytes() == b"EXISTING USER FILE"
    assert not _stages(tmp_path)


def test_missing_staged_file_is_refused_without_published_output(tmp_path):
    target = tmp_path / "missing.tsv"
    with pytest.raises(publication.VerifiedReportPublicationError, match="missing"):
        with publication.staged_report_outputs((target,)) as (_stage,):
            pass
    assert not target.exists()
    assert not _stages(tmp_path)


def test_duplicate_or_symlink_output_collisions_are_not_published(tmp_path):
    target = tmp_path / "collision.tsv"
    with pytest.raises(publication.VerifiedReportPublicationError, match="collide"):
        with publication.staged_report_outputs((target, target)):
            raise AssertionError("Must not yield")
    target.write_bytes(b"USER ORIGINAL")
    alias = tmp_path / "alias.tsv"
    try:
        alias.symlink_to(target)
    except (OSError, NotImplementedError):
        return
    with pytest.raises(publication.VerifiedReportPublicationError, match="already exists|collide"):
        with publication.staged_report_outputs((alias,)):
            raise AssertionError("Must not yield")
    assert target.read_bytes() == b"USER ORIGINAL"


def test_double_report_publish_rolls_back_only_own_new_first_file(tmp_path, monkeypatch):
    main = tmp_path / "therapist_main.tsv"
    trend = tmp_path / "therapist_trend.tsv"
    original_link = publication.os.link
    calls = []

    def fail_second(src, dst):
        calls.append(Path(dst))
        if len(calls) == 2:
            raise OSError("FICTIONAL SECOND OUTPUT FAILURE")
        return original_link(src, dst)

    monkeypatch.setattr(publication.os, "link", fail_second)
    with pytest.raises(OSError, match="SECOND OUTPUT FAILURE"):
        with publication.staged_report_outputs((main, trend)) as stages:
            stages[0].write_bytes(b"VERIFIED MAIN")
            stages[1].write_bytes(b"VERIFIED TREND")
    assert calls == [main, trend]
    assert not main.exists() and not trend.exists()
    assert not _stages(tmp_path)


def test_competing_file_created_after_preflight_is_preserved(tmp_path, monkeypatch):
    target = tmp_path / "patient_report.tsv"
    original_link = publication.os.link
    competing = b"EXTERNAL USER DATA"

    def create_competitor(src, dst):
        Path(dst).write_bytes(competing)
        return original_link(src, dst)

    monkeypatch.setattr(publication.os, "link", create_competitor)
    with pytest.raises(FileExistsError):
        with publication.staged_report_outputs((target,)) as (stage,):
            stage.write_bytes(b"VERIFIED REPORT")
    assert target.read_bytes() == competing
    assert not _stages(tmp_path)


def test_multifile_success_supports_two_different_output_directories(tmp_path):
    primary = tmp_path / "one" / "main.tsv"
    secondary = tmp_path / "two" / "trend.tsv"
    with publication.staged_report_outputs((primary, secondary)) as stages:
        assert all(stage.parent != target.parent for stage, target in zip(stages, (primary, secondary)))
        for stage, name in zip(stages, (b"MAIN", b"TREND")):
            stage.write_bytes(name)
    assert primary.read_bytes() == b"MAIN"
    assert secondary.read_bytes() == b"TREND"
    assert not _stages(primary.parent)
    assert not _stages(secondary.parent)
