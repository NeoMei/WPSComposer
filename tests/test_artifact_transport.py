from __future__ import annotations

import os
from pathlib import Path
import pickle
import subprocess
import sys
from types import MappingProxyType
import threading
import time
import zipfile

import pytest

from skills.WPSComposer.scripts import artifact_transport
from skills.WPSComposer.scripts.artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    ValidatorSpec,
    copy_file_before_deadline,
    publish_artifact,
    publish_artifact_group,
    snapshot_artifact_state,
    validate_office_package,
    validate_before_deadline,
    validate_pdf,
)
from tests._pdf_fixture import write_minimal_pdf


def _write_pdf(path: Path, payload: bytes = b"x" * 2048) -> Path:
    return write_minimal_pdf(path, payload)


def _write_package(path: Path, member: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr(member, f"<root><data>{'x' * 2048}</data></root>")
    return path


def _accept_validator(_path: Path) -> None:
    return None


def _blocked_validator(_path: Path) -> None:
    time.sleep(1)


def test_validate_pdf_requires_signature_and_minimum_size(tmp_path: Path):
    validate_pdf(_write_pdf(tmp_path / "valid.pdf"))

    bad = tmp_path / "bad.pdf"
    bad.write_bytes(b"not-a-pdf" + b"x" * 2048)
    with pytest.raises(ArtifactValidationError, match="Invalid PDF signature"):
        validate_pdf(bad)


def test_validate_pdf_rejects_signature_prefixed_garbage(tmp_path: Path):
    corrupt = tmp_path / "corrupt.pdf"
    corrupt.write_bytes(b"%PDF-1.7\n" + b"not a PDF object graph" * 100)

    with pytest.raises(ArtifactValidationError, match="Invalid PDF structure"):
        validate_pdf(corrupt)


@pytest.mark.parametrize(
    ("format_name", "member"),
    [
        ("docx", "word/document.xml"),
        ("xlsx", "xl/workbook.xml"),
        ("pptx", "ppt/presentation.xml"),
    ],
)
def test_validate_office_package_requires_expected_member(
    tmp_path: Path, format_name: str, member: str
):
    valid = _write_package(tmp_path / f"valid.{format_name}", member)
    validate_office_package(valid, format_name)

    incomplete = _write_package(
        tmp_path / f"incomplete.{format_name}", "custom/missing.xml"
    )
    with pytest.raises(ArtifactValidationError, match=member):
        validate_office_package(incomplete, format_name)


def test_validate_office_package_rejects_malformed_defining_xml(tmp_path: Path):
    package_path = tmp_path / "corrupt.docx"
    with zipfile.ZipFile(package_path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr("word/document.xml", "<broken>" + "x" * 2048)

    with pytest.raises(ArtifactValidationError, match="malformed XML"):
        validate_office_package(package_path, "docx")


def test_publish_validates_staged_temporary_and_final_copies(
    tmp_path: Path,
):
    staged = _write_pdf(tmp_path / "stage" / "result.pdf")
    destination = tmp_path / "output" / "result.pdf"
    validated: list[Path] = []

    def validator(path: Path) -> None:
        target = Path(path).resolve()
        validated.append(target)
        validate_pdf(target)

    result = publish_artifact(
        staged,
        destination,
        overwrite=False,
        validator=validator,
    )

    assert result == destination.resolve()
    assert validated[0] == staged.resolve()
    assert validated[-1] == destination.resolve()
    assert len(validated) == 3
    assert validated[1].parent == destination.parent.resolve()
    assert validated[1].name.startswith(".wpscomposer-")
    assert not validated[1].exists()
    assert result.read_bytes() == staged.read_bytes()


def test_conditional_publish_preserves_destination_changed_after_preflight(
    tmp_path: Path,
):
    staged = _write_pdf(tmp_path / "stage" / "result.pdf", b"edited" * 400)
    destination = _write_pdf(tmp_path / "result.pdf", b"original" * 400)
    expected = snapshot_artifact_state(destination)
    concurrent = _write_pdf(tmp_path / "concurrent.pdf", b"concurrent" * 400).read_bytes()
    destination.write_bytes(concurrent)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=validate_pdf,
            expected_destination=expected,
        )

    assert caught.value.code == "ARTIFACT_DESTINATION_CHANGED"
    assert destination.read_bytes() == concurrent


def test_conditional_publish_keeps_an_unchanged_logical_symlink(tmp_path: Path):
    staged = _write_pdf(tmp_path / "stage.pdf", b"edited" * 400)
    target = _write_pdf(tmp_path / "target.pdf", b"original" * 400)
    logical = tmp_path / "logical.pdf"
    logical.symlink_to(target)
    expected = snapshot_artifact_state(logical)

    result = publish_artifact(
        staged,
        logical,
        overwrite=True,
        validator=validate_pdf,
        expected_destination=expected,
    )

    assert result == target.resolve()
    assert logical.is_symlink()
    assert logical.resolve().read_bytes() == staged.read_bytes()


@pytest.mark.parametrize("replacement", ["deleted", "symlink"])
def test_conditional_publish_rejects_deleted_or_retargeted_logical_destination(
    tmp_path: Path, replacement: str,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"edited" * 400)
    logical = _write_pdf(tmp_path / "logical.pdf", b"original" * 400)
    expected = snapshot_artifact_state(logical)
    if replacement == "deleted":
        logical.unlink()
    else:
        other = _write_pdf(tmp_path / "other.pdf", b"other" * 400)
        logical.unlink()
        logical.symlink_to(other)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            expected.resolved_path,
            overwrite=True,
            validator=validate_pdf,
            expected_destination=expected,
        )

    assert caught.value.code == "ARTIFACT_DESTINATION_CHANGED"
    if replacement == "symlink":
        assert logical.is_symlink()


def test_conditional_group_checks_every_destination_before_first_replacement(
    tmp_path: Path,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"first-new" * 400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"second-new" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"first-old" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"second-old" * 400)
    first_before = first.read_bytes()
    first_expected = snapshot_artifact_state(first)
    second_expected = snapshot_artifact_state(second)
    concurrent = _write_pdf(tmp_path / "concurrent.pdf", b"second-concurrent" * 300).read_bytes()
    second.write_bytes(concurrent)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact_group(
            [
                (first_stage, first, True, validate_pdf, first_expected),
                (second_stage, second, True, validate_pdf, second_expected),
            ]
        )

    assert caught.value.code == "ARTIFACT_DESTINATION_CHANGED"
    assert first.read_bytes() == first_before
    assert second.read_bytes() == concurrent


def test_single_publish_rolls_back_when_deadline_expires_after_replace(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 800)
    destination = _write_pdf(tmp_path / "destination.pdf", b"old" * 800)
    original = destination.read_bytes()
    real_replace = os.replace
    expired = False

    def expire_after_replace(source, target):
        nonlocal expired
        real_replace(source, target)
        if Path(target) == destination:
            expired = True

    def validator(path):
        validate_pdf(path)
        if Path(path) == destination and expired:
            raise TimeoutError("deadline expired after replacement")

    monkeypatch.setattr(os, "replace", expire_after_replace)
    with pytest.raises(ArtifactTransportError, match="deadline expired"):
        publish_artifact(staged, destination, overwrite=True, validator=validator)

    assert destination.read_bytes() == original


def test_group_rolls_back_when_final_validation_times_out_after_replace(
    tmp_path: Path,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"new-first" * 400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"new-second" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"old-first" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"old-second" * 400)
    original = (first.read_bytes(), second.read_bytes())

    def fail_final(path):
        validate_pdf(path)
        if Path(path) == second:
            raise TimeoutError("late group validation timeout")

    with pytest.raises(ArtifactTransportError, match="late group"):
        publish_artifact_group([
            (first_stage, first, True, validate_pdf),
            (second_stage, second, True, fail_final),
        ])

    assert first.read_bytes() == original[0]
    assert second.read_bytes() == original[1]


def test_failed_final_validation_preserves_a_concurrent_replacement(
    tmp_path: Path,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 800)
    destination = _write_pdf(tmp_path / "destination.pdf", b"old" * 800)
    concurrent = _write_pdf(tmp_path / "concurrent.pdf", b"concurrent" * 400)
    concurrent_bytes = concurrent.read_bytes()

    def replace_during_final_validation(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == destination:
            os.replace(concurrent, destination)
            raise ArtifactValidationError("late validation failure")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=replace_during_final_validation,
        )

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.read_bytes() == concurrent_bytes
    assert list(tmp_path.glob(".wpscomposer-backup-*.tmp"))


def test_successful_final_validation_rejects_a_foreign_valid_replacement(
    tmp_path: Path,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 800)
    destination = _write_pdf(tmp_path / "destination.pdf", b"old" * 800)
    old_bytes = destination.read_bytes()
    concurrent = _write_pdf(tmp_path / "concurrent.pdf", b"concurrent" * 400)
    concurrent_bytes = concurrent.read_bytes()

    def replace_during_final_validation(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == destination:
            os.replace(concurrent, destination)
            validate_pdf(destination)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=replace_during_final_validation,
        )

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.read_bytes() == concurrent_bytes
    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == old_bytes


def test_successful_final_validation_checks_hash_when_inode_and_stat_match(
    tmp_path: Path,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"A" * 2400)
    destination = _write_pdf(tmp_path / "destination.pdf", b"old" * 800)
    old_bytes = destination.read_bytes()
    foreign = _write_pdf(tmp_path / "foreign.pdf", b"B" * 2400)
    foreign_bytes = foreign.read_bytes()
    assert len(foreign_bytes) == len(staged.read_bytes())

    def mutate_bytes_but_restore_stat(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == destination:
            before = path.stat()
            path.write_bytes(foreign_bytes)
            os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
            after = path.stat()
            assert (after.st_ino, after.st_size, after.st_mtime_ns) == (
                before.st_ino, before.st_size, before.st_mtime_ns
            )
            validate_pdf(path)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=mutate_bytes_but_restore_stat,
        )

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.read_bytes() == foreign_bytes
    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == old_bytes


def test_failed_final_validation_checks_content_before_rollback(tmp_path: Path):
    staged = _write_pdf(tmp_path / "stage.pdf", b"A" * 2400)
    destination = _write_pdf(tmp_path / "destination.pdf", b"old" * 800)
    old_bytes = destination.read_bytes()
    foreign_bytes = _write_pdf(tmp_path / "foreign.pdf", b"B" * 2400).read_bytes()
    assert len(foreign_bytes) == len(staged.read_bytes())

    def mutate_in_place_then_fail(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == destination:
            before = path.stat()
            path.write_bytes(foreign_bytes)
            os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
            assert path.stat().st_ino == before.st_ino
            raise ArtifactValidationError("validator rejected after concurrent write")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=mutate_in_place_then_fail,
        )

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.read_bytes() == foreign_bytes
    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == old_bytes


def test_failed_final_validation_preserves_output_when_ownership_hash_times_out(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 800)
    destination = _write_pdf(tmp_path / "destination.pdf", b"old" * 800)
    old_bytes = destination.read_bytes()
    published_bytes = staged.read_bytes()
    validator_failed = False
    real_snapshot = artifact_transport.snapshot_artifact_state

    def fail_final(path: Path) -> None:
        nonlocal validator_failed
        validate_pdf(path)
        if Path(path) == destination:
            validator_failed = True
            raise ArtifactValidationError("late failure")

    def time_out_ownership_snapshot(path: Path, *, deadline=None):
        if validator_failed and Path(path) == destination:
            raise TimeoutError("stable ownership snapshot timed out")
        return real_snapshot(path, deadline=deadline)

    monkeypatch.setattr(
        artifact_transport, "snapshot_artifact_state", time_out_ownership_snapshot
    )
    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged, destination, overwrite=True, validator=fail_final
        )

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.read_bytes() == published_bytes
    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == old_bytes


def test_final_validation_rejects_a_symlink_to_the_published_inode(tmp_path: Path):
    staged = tmp_path / "stage.bin"
    staged.write_bytes(b"published bytes")
    destination = tmp_path / "destination.bin"
    destination.write_bytes(b"previous bytes")
    old_bytes = destination.read_bytes()
    moved = tmp_path / "externally-moved.bin"

    def replace_leaf_with_symlink(path: Path) -> None:
        if Path(path) == destination:
            destination.rename(moved)
            destination.symlink_to(moved)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=replace_leaf_with_symlink,
        )

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.is_symlink() and destination.resolve() == moved
    assert moved.read_bytes() == staged.read_bytes()
    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == old_bytes


def test_single_link_cleanup_failure_rolls_back_new_destination(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
):
    source = tmp_path / "source.bin"
    source.write_bytes(b"new-data")
    target = tmp_path / "target.bin"
    real_unlink = Path.unlink
    failures = []

    def fail_first_temporary_cleanup(self, *args, **kwargs):
        if self.name.startswith(".wpscomposer-") and target.exists() and not failures:
            failures.append(True)
            raise OSError("cleanup denied")
        return real_unlink(self, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", fail_first_temporary_cleanup)

    with pytest.raises(ArtifactTransportError):
        publish_artifact(
            source, target, overwrite=False, validator=lambda _path: None
        )

    assert not target.exists()


def test_group_rollback_preserves_a_concurrently_replaced_member(
    tmp_path: Path,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"new-first" * 400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"new-second" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"old-first" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"old-second" * 400)
    second_before = second.read_bytes()
    concurrent = _write_pdf(tmp_path / "concurrent.pdf", b"concurrent" * 400)
    concurrent_bytes = concurrent.read_bytes()

    def replace_first_during_final_validation(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == first:
            os.replace(concurrent, first)
            raise ArtifactValidationError("late group validation failure")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact_group([
            (first_stage, first, True, replace_first_during_final_validation),
            (second_stage, second, True, validate_pdf),
        ])

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert first.read_bytes() == concurrent_bytes
    assert second.read_bytes() == second_before
    assert list(tmp_path.glob(".wpscomposer-group-backup-*.tmp"))


def test_group_final_pass_detects_a_later_validator_replacing_an_earlier_member(
    tmp_path: Path,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"new-first" * 400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"new-second" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"old-first" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"old-second" * 400)
    first_before = first.read_bytes()
    second_before = second.read_bytes()
    concurrent = _write_pdf(tmp_path / "concurrent.pdf", b"concurrent" * 400)
    concurrent_bytes = concurrent.read_bytes()

    def replace_first_while_validating_second(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == second:
            os.replace(concurrent, first)
            validate_pdf(first)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact_group([
            (first_stage, first, True, validate_pdf),
            (second_stage, second, True, replace_first_while_validating_second),
        ])

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert first.read_bytes() == concurrent_bytes
    assert second.read_bytes() == second_before
    backups = list(tmp_path.glob(".wpscomposer-group-backup-*.tmp"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == first_before


def test_group_validator_error_checks_every_published_member_before_rollback(
    tmp_path: Path,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"A" * 2400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"new-second" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"old-first" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"old-second" * 400)
    first_old = first.read_bytes()
    second_old = second.read_bytes()
    foreign_bytes = _write_pdf(tmp_path / "foreign.pdf", b"B" * 2400).read_bytes()
    assert len(foreign_bytes) == len(first_stage.read_bytes())

    def mutate_first_then_fail(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == second:
            before = first.stat()
            first.write_bytes(foreign_bytes)
            os.utime(first, ns=(before.st_atime_ns, before.st_mtime_ns))
            assert first.stat().st_ino == before.st_ino
            raise ArtifactValidationError("later member rejected")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact_group([
            (first_stage, first, True, validate_pdf),
            (second_stage, second, True, mutate_first_then_fail),
        ])

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert first.read_bytes() == foreign_bytes
    assert second.read_bytes() == second_old
    backups = list(tmp_path.glob(".wpscomposer-group-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == first_old


def test_group_publish_failure_content_checks_an_earlier_visible_member(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"A" * 2400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"new-second" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"old-first" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"old-second" * 400)
    first_old = first.read_bytes()
    second_old = second.read_bytes()
    foreign_bytes = _write_pdf(tmp_path / "foreign.pdf", b"B" * 2400).read_bytes()
    real_replace = os.replace

    def fail_second_publication(source, target):
        if Path(target) == second and Path(source).name.startswith(
            ".wpscomposer-group-publish-"
        ):
            before = first.stat()
            first.write_bytes(foreign_bytes)
            os.utime(first, ns=(before.st_atime_ns, before.st_mtime_ns))
            raise OSError("second publication failed")
        return real_replace(source, target)

    monkeypatch.setattr(os, "replace", fail_second_publication)
    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact_group([
            (first_stage, first, True, validate_pdf),
            (second_stage, second, True, validate_pdf),
        ])

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert first.read_bytes() == foreign_bytes
    assert second.read_bytes() == second_old
    backups = list(tmp_path.glob(".wpscomposer-group-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == first_old


def test_group_validator_error_preserves_symlink_to_published_inode(tmp_path: Path):
    staged = tmp_path / "stage.bin"
    staged.write_bytes(b"published bytes")
    destination = tmp_path / "destination.bin"
    destination.write_bytes(b"previous bytes")
    old_bytes = destination.read_bytes()
    moved = tmp_path / "externally-moved.bin"

    def replace_leaf_then_fail(path: Path) -> None:
        if Path(path) == destination:
            destination.rename(moved)
            destination.symlink_to(moved)
            raise ArtifactValidationError("path binding changed")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact_group([
            (staged, destination, True, replace_leaf_then_fail),
        ])

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert destination.is_symlink() and destination.resolve() == moved
    assert moved.read_bytes() == staged.read_bytes()
    backups = list(tmp_path.glob(".wpscomposer-group-backup-*.tmp"))
    assert len(backups) == 1 and backups[0].read_bytes() == old_bytes


@pytest.mark.parametrize("failure", [KeyboardInterrupt(), SystemExit(17)])
def test_group_final_validation_rethrows_control_exception_after_safe_rollback(
    tmp_path: Path, failure: BaseException,
):
    first_stage = _write_pdf(tmp_path / "first-stage.pdf", b"new-first" * 400)
    second_stage = _write_pdf(tmp_path / "second-stage.pdf", b"new-second" * 400)
    first = _write_pdf(tmp_path / "first.pdf", b"old-first" * 400)
    second = _write_pdf(tmp_path / "second.pdf", b"old-second" * 400)
    original = (first.read_bytes(), second.read_bytes())

    def interrupt_second(path: Path) -> None:
        validate_pdf(path)
        if Path(path) == second:
            raise failure

    with pytest.raises(type(failure)) as caught:
        publish_artifact_group([
            (first_stage, first, True, validate_pdf),
            (second_stage, second, True, interrupt_second),
        ])

    assert caught.value is failure
    assert (first.read_bytes(), second.read_bytes()) == original


def test_snapshot_checks_deadline_after_slow_digest_work(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
):
    path = _write_pdf(tmp_path / "artifact.pdf")
    real_sha256 = artifact_transport.hashlib.sha256

    class SlowDigest:
        def __init__(self):
            self._digest = real_sha256()

        def update(self, data):
            time.sleep(0.02)
            self._digest.update(data)

        def hexdigest(self):
            return self._digest.hexdigest()

    monkeypatch.setattr(artifact_transport.hashlib, "sha256", SlowDigest)

    with pytest.raises(TimeoutError):
        snapshot_artifact_state(path, deadline=time.monotonic() + 0.005)


def test_publish_refuses_existing_output_without_overwrite(tmp_path: Path):
    staged = _write_pdf(tmp_path / "stage.pdf")
    destination = _write_pdf(tmp_path / "result.pdf", b"old" * 1024)
    original = destination.read_bytes()

    with pytest.raises(FileExistsError, match="Output already exists"):
        publish_artifact(
            staged,
            destination,
            overwrite=False,
            validator=validate_pdf,
        )

    assert destination.read_bytes() == original


def test_publish_overwrites_atomically_when_requested(tmp_path: Path):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 1024)
    destination = _write_pdf(tmp_path / "result.pdf", b"old" * 1024)

    result = publish_artifact(
        staged,
        destination,
        overwrite=True,
        validator=validate_pdf,
    )

    assert result.read_bytes() == staged.read_bytes()
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))


def test_group_publish_rolls_back_all_outputs_when_later_publish_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    first_stage = _write_pdf(tmp_path / "stage" / "first.pdf", b"new-first")
    second_stage = _write_pdf(tmp_path / "stage" / "second.pdf", b"new-second")
    first = _write_pdf(tmp_path / "out" / "first.pdf", b"old-first")
    second = _write_pdf(tmp_path / "out" / "second.pdf", b"old-second")
    originals = (first.read_bytes(), second.read_bytes())
    real_replace = os.replace
    publish_count = 0

    def fail_second_publish(source, target):
        nonlocal publish_count
        source_path = Path(source)
        if source_path.name.startswith(".wpscomposer-group-"):
            publish_count += 1
            if publish_count == 2:
                raise OSError("simulated second publish failure")
        return real_replace(source, target)

    monkeypatch.setattr(os, "replace", fail_second_publish)

    with pytest.raises(ArtifactTransportError, match="second publish failure"):
        artifact_transport.publish_artifact_group([
            (first_stage, first, True, validate_pdf),
            (second_stage, second, True, validate_pdf),
        ])

    assert (first.read_bytes(), second.read_bytes()) == originals
    assert not list((tmp_path / "out").glob(".wpscomposer-group-*.tmp"))
    assert not list((tmp_path / "out").glob(".wpscomposer-group-backup-*.tmp"))


def test_group_publish_cleans_partial_backup_when_backup_copy_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "stage" / "result.pdf", b"new")
    destination = _write_pdf(tmp_path / "out" / "result.pdf", b"secret-old")
    original = destination.read_bytes()
    real_copy = artifact_transport.copy_stream_before_deadline

    def fail_backup_copy(incoming, outgoing, deadline):
        if Path(outgoing.name).name.startswith(".wpscomposer-group-backup-"):
            outgoing.write(b"secret-prefix")
            raise OSError("simulated backup copy failure")
        return real_copy(incoming, outgoing, deadline)

    monkeypatch.setattr(
        artifact_transport, "copy_stream_before_deadline", fail_backup_copy
    )

    with pytest.raises(ArtifactTransportError) as caught:
        artifact_transport.publish_artifact_group([
            (staged, destination, True, validate_pdf),
        ])

    assert caught.value.code == "ARTIFACT_PUBLISH_FAILED"
    assert destination.read_bytes() == original
    assert not list((tmp_path / "out").glob(".wpscomposer-group-*.tmp"))
    assert not list((tmp_path / "out").glob(".wpscomposer-group-backup-*.tmp"))


def test_group_publish_reports_partial_backup_when_cleanup_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "stage" / "result.pdf", b"new")
    destination = _write_pdf(tmp_path / "out" / "result.pdf", b"secret-old")
    real_copy = artifact_transport.copy_stream_before_deadline
    real_unlink = Path.unlink
    retained: list[Path] = []

    def fail_backup_copy(incoming, outgoing, deadline):
        if Path(outgoing.name).name.startswith(".wpscomposer-group-backup-"):
            outgoing.write(b"secret-prefix")
            raise OSError("simulated backup copy failure")
        return real_copy(incoming, outgoing, deadline)

    def fail_backup_cleanup(path, *args, **kwargs):
        if path.name.startswith(".wpscomposer-group-backup-"):
            retained.append(path.resolve())
            raise PermissionError("simulated cleanup failure")
        return real_unlink(path, *args, **kwargs)

    monkeypatch.setattr(
        artifact_transport, "copy_stream_before_deadline", fail_backup_copy
    )
    monkeypatch.setattr(Path, "unlink", fail_backup_cleanup)

    with pytest.raises(ArtifactTransportError) as caught:
        artifact_transport.publish_artifact_group([
            (staged, destination, True, validate_pdf),
        ])

    assert caught.value.code == "ARTIFACT_CLEANUP_FAILED"
    assert len(retained) == 1
    assert str(retained[0]) in str(caught.value)
    assert retained[0].read_bytes() == b"secret-prefix"


def test_group_publish_rolls_back_target_when_post_link_cleanup_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    first_stage = _write_pdf(tmp_path / "stage" / "first.pdf", b"new-first")
    second_stage = _write_pdf(tmp_path / "stage" / "second.pdf", b"new-second")
    first = tmp_path / "out" / "first.pdf"
    second = tmp_path / "out" / "second.pdf"
    real_link = os.link
    real_unlink = Path.unlink
    second_local: Path | None = None
    cleanup_failed = False

    def observe_link(source, target):
        nonlocal second_local
        result = real_link(source, target)
        if Path(target).resolve() == second.resolve():
            second_local = Path(source).resolve()
        return result

    def fail_second_local_cleanup(path, *args, **kwargs):
        nonlocal cleanup_failed
        if (
            second_local is not None
            and path.resolve() == second_local
            and not cleanup_failed
        ):
            cleanup_failed = True
            raise OSError("simulated post-link cleanup failure")
        return real_unlink(path, *args, **kwargs)

    monkeypatch.setattr(os, "link", observe_link)
    monkeypatch.setattr(Path, "unlink", fail_second_local_cleanup)

    with pytest.raises(ArtifactTransportError, match="post-link cleanup failure"):
        artifact_transport.publish_artifact_group([
            (first_stage, first, False, validate_pdf),
            (second_stage, second, False, validate_pdf),
        ])

    assert cleanup_failed
    assert not first.exists()
    assert not second.exists()
    assert not list((tmp_path / "out").glob(".wpscomposer-group-*.tmp"))


def test_group_publish_reports_target_when_post_link_rollback_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "stage" / "result.pdf", b"new")
    target = tmp_path / "out" / "result.pdf"
    real_link = os.link
    real_unlink = Path.unlink
    published_local: Path | None = None
    cleanup_failed = False

    def observe_link(source, destination):
        nonlocal published_local
        result = real_link(source, destination)
        published_local = Path(source).resolve()
        return result

    def fail_cleanup_and_rollback(path, *args, **kwargs):
        nonlocal cleanup_failed
        resolved = path.resolve()
        if resolved == published_local and not cleanup_failed:
            cleanup_failed = True
            raise OSError("simulated post-link cleanup failure")
        if resolved == target.resolve():
            raise OSError("simulated rollback failure")
        return real_unlink(path, *args, **kwargs)

    monkeypatch.setattr(os, "link", observe_link)
    monkeypatch.setattr(Path, "unlink", fail_cleanup_and_rollback)

    with pytest.raises(ArtifactTransportError) as caught:
        artifact_transport.publish_artifact_group([
            (staged, target, False, validate_pdf),
        ])

    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert str(target.resolve()) in str(caught.value)
    assert target.exists()


def test_publish_validation_failure_leaves_no_partial_destination(
    tmp_path: Path,
):
    staged = tmp_path / "bad.pdf"
    staged.write_bytes(b"broken")
    destination = tmp_path / "result.pdf"

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=False,
            validator=validate_pdf,
        )

    assert caught.value.code == "STAGED_ARTIFACT_INVALID"
    assert not destination.exists()
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))


def test_publish_uses_os_replace_in_destination_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "staging" / "result.pdf")
    destination = tmp_path / "external" / "result.pdf"
    calls: list[tuple[Path, Path]] = []
    original_replace = os.replace

    def observed_replace(source, target):
        calls.append((Path(source).resolve(), Path(target).resolve()))
        original_replace(source, target)

    monkeypatch.setattr(os, "replace", observed_replace)

    publish_artifact(
        staged,
        destination,
        overwrite=True,
        validator=validate_pdf,
    )

    assert len(calls) == 1
    assert calls[0][0].parent == destination.parent.resolve()
    assert calls[0][1] == destination.resolve()


def test_publish_does_not_clobber_target_created_during_no_overwrite_race(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"staged" * 1024)
    destination = tmp_path / "result.pdf"
    intruder_path = write_minimal_pdf(tmp_path / "intruder.pdf", b"intruder")
    intruder = intruder_path.read_bytes()
    intruder_path.unlink()
    original_link = os.link

    def raced_link(source, target):
        Path(target).write_bytes(intruder)
        return original_link(source, target)

    monkeypatch.setattr(os, "link", raced_link)

    with pytest.raises(FileExistsError):
        publish_artifact(
            staged,
            destination,
            overwrite=False,
            validator=validate_pdf,
        )

    assert destination.read_bytes() == intruder
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))


def test_publish_restores_existing_output_when_final_validation_fails(tmp_path):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new artifact")
    destination = _write_pdf(tmp_path / "result.pdf", b"old artifact")
    original = destination.read_bytes()
    replacement = staged.read_bytes()

    def fail_only_after_replace(path):
        validate_pdf(path)
        target = Path(path).resolve()
        if target == destination.resolve() and target.read_bytes() == replacement:
            raise ArtifactValidationError("simulated final validation failure")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=fail_only_after_replace,
        )

    assert caught.value.code == "FINAL_ARTIFACT_INVALID"
    assert destination.read_bytes() == original


def test_publish_restores_existing_output_for_unexpected_validator_exception(tmp_path):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new artifact")
    destination = _write_pdf(tmp_path / "result.pdf", b"old artifact")
    original = destination.read_bytes()
    replacement = staged.read_bytes()

    def fail_only_after_replace(path):
        target = Path(path).resolve()
        if target == destination.resolve() and target.read_bytes() == replacement:
            raise RuntimeError("parser crashed")

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=fail_only_after_replace,
        )

    assert caught.value.code == "FINAL_ARTIFACT_INVALID"
    assert destination.read_bytes() == original
    assert not list(tmp_path.glob(".wpscomposer-backup-*.tmp"))


@pytest.mark.parametrize("failure", [KeyboardInterrupt(), SystemExit(17)])
def test_publish_restores_existing_output_and_rethrows_control_exception(
    tmp_path: Path, failure: BaseException
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new artifact")
    destination = _write_pdf(tmp_path / "result.pdf", b"old artifact")
    original = destination.read_bytes()
    replacement = staged.read_bytes()

    def interrupt_only_after_replace(path):
        target = Path(path).resolve()
        if target == destination.resolve() and target.read_bytes() == replacement:
            raise failure

    with pytest.raises(type(failure)) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=interrupt_only_after_replace,
        )

    assert caught.value is failure
    assert destination.read_bytes() == original
    assert not list(tmp_path.glob(".wpscomposer-backup-*.tmp"))


@pytest.mark.parametrize("failure", [KeyboardInterrupt(), SystemExit(17)])
def test_publish_retains_backup_and_interrupt_semantics_when_rollback_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    failure: BaseException,
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new artifact")
    destination = _write_pdf(tmp_path / "result.pdf", b"old artifact")
    replacement = staged.read_bytes()
    real_replace = os.replace
    calls = 0

    def fail_restore(source, target):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("restore blocked")
        return real_replace(source, target)

    def interrupt_only_after_replace(path):
        target = Path(path).resolve()
        if target == destination.resolve() and target.read_bytes() == replacement:
            raise failure

    monkeypatch.setattr(os, "replace", fail_restore)
    with pytest.raises(type(failure)) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=interrupt_only_after_replace,
        )

    assert caught.value is failure
    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert len(backups) == 1
    rollback_error = caught.value.__cause__
    assert isinstance(rollback_error, ArtifactTransportError)
    assert rollback_error.code == "ARTIFACT_ROLLBACK_FAILED"
    assert str(backups[0]) in str(rollback_error)
    assert isinstance(rollback_error.__cause__, OSError)


def test_publish_retains_and_reports_backup_when_rollback_fails(
    tmp_path, monkeypatch
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new artifact")
    destination = _write_pdf(tmp_path / "result.pdf", b"old artifact")
    replacement = staged.read_bytes()
    real_replace = os.replace
    calls = 0

    def fail_restore(source, target):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("restore blocked")
        return real_replace(source, target)

    def fail_only_after_replace(path):
        target = Path(path).resolve()
        if target == destination.resolve() and target.read_bytes() == replacement:
            raise RuntimeError("parser crashed")

    monkeypatch.setattr(os, "replace", fail_restore)
    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=fail_only_after_replace,
        )

    backups = list(tmp_path.glob(".wpscomposer-backup-*.tmp"))
    assert caught.value.code == "ARTIFACT_ROLLBACK_FAILED"
    assert len(backups) == 1
    assert str(backups[0]) in str(caught.value)


def test_publish_chunk_copy_stops_at_deadline_and_preserves_target(
    tmp_path, monkeypatch
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 1024 * 1024)
    destination = _write_pdf(tmp_path / "result.pdf", b"old artifact")
    original = destination.read_bytes()
    now = [10.0]

    def ticking_clock():
        value = now[0]
        now[0] += 0.025
        return value

    monkeypatch.setattr(artifact_transport.time, "monotonic", ticking_clock)

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=True,
            validator=lambda path: None,
            deadline=10.2,
        )

    assert caught.value.code == "ARTIFACT_PUBLISH_FAILED"
    assert destination.read_bytes() == original
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))


def test_read_only_validator_is_bounded_by_absolute_deadline(tmp_path):
    started = time.monotonic()

    with pytest.raises(TimeoutError, match="validation deadline"):
        validate_before_deadline(
            _blocked_validator,
            tmp_path / "artifact.pdf",
            deadline=time.monotonic() + 0.02,
        )

    assert time.monotonic() - started < 0.15


# /dev/fd enumeration is POSIX-only
@pytest.mark.skipif(
    os.name != "posix", reason="/dev/fd fd-growth check is POSIX-only"
)
def test_repeated_validator_timeouts_leave_no_worker_or_fd_growth(tmp_path):
    before_threads = {
        thread.ident for thread in threading.enumerate() if thread.is_alive()
    }
    before_fds = len(list(Path("/dev/fd").iterdir()))

    for _ in range(4):
        with pytest.raises(TimeoutError, match="validation deadline"):
            validate_before_deadline(
                _blocked_validator,
                tmp_path / "artifact.pdf",
                deadline=time.monotonic() + 0.01,
            )

    time.sleep(0.02)
    after_threads = {
        thread.ident for thread in threading.enumerate() if thread.is_alive()
    }
    after_fds = len(list(Path("/dev/fd").iterdir()))
    assert after_threads == before_threads
    assert after_fds <= before_fds + 1


def test_deadline_validator_uses_spawn_safe_worker_context(tmp_path, monkeypatch):
    real_get_context = artifact_transport.multiprocessing.get_context
    requested = []

    def observed_get_context(method):
        requested.append(method)
        return real_get_context(method)

    monkeypatch.setattr(
        artifact_transport.multiprocessing, "get_context", observed_get_context
    )

    validate_before_deadline(
        _accept_validator,
        tmp_path / "artifact.pdf",
        deadline=time.monotonic() + 2,
    )

    assert requested == ["spawn"]


def test_deadline_validator_accepts_spawn_safe_explicit_spec(tmp_path):
    package = _write_package(tmp_path / "artifact.docx", "word/document.xml")
    spec = ValidatorSpec.from_callable(validate_office_package, "docx")

    validate_before_deadline(
        spec,
        package,
        deadline=time.monotonic() + 2,
    )


def test_deadline_validator_works_from_python_dash_c(tmp_path):
    pdf = _write_pdf(tmp_path / "interactive.pdf")
    script = (
        "import time; "
        "from pathlib import Path; "
        "from skills.WPSComposer.scripts.artifact_transport import "
        "ValidatorSpec, validate_before_deadline, validate_pdf; "
        f"validate_before_deadline(ValidatorSpec.from_callable(validate_pdf), "
        f"Path({str(pdf)!r}), time.monotonic() + 10)"
    )
    result = subprocess.run(
        [sys.executable, "-c", script],
        cwd=Path(__file__).resolve().parents[1],
        capture_output=True,
        text=True,
        timeout=15,
    )
    assert result.returncode == 0, result.stderr


def test_deadline_validator_works_from_python_stdin(tmp_path):
    pdf = _write_pdf(tmp_path / "stdin.pdf")
    script = (
        "import time\n"
        "from pathlib import Path\n"
        "from skills.WPSComposer.scripts.artifact_transport import "
        "ValidatorSpec, validate_before_deadline, validate_pdf\n"
        f"validate_before_deadline(ValidatorSpec.from_callable(validate_pdf), "
        f"Path({str(pdf)!r}), time.monotonic() + 10)\n"
    )
    result = subprocess.run(
        [sys.executable, "-"],
        input=script,
        cwd=Path(__file__).resolve().parents[1],
        capture_output=True,
        text=True,
        timeout=15,
    )
    assert result.returncode == 0, result.stderr


def test_validator_spec_rejects_non_plain_arguments():
    with pytest.raises(TypeError, match="plain serializable"):
        ValidatorSpec.from_callable(
            validate_office_package,
            MappingProxyType({"format": "docx"}),
        )


def test_deadline_validator_rejects_unpicklable_callable_without_child_leak(
    tmp_path,
):
    captured = MappingProxyType({"format": "pdf"})

    def unpicklable(_path):
        return captured

    before = {child.pid for child in artifact_transport.multiprocessing.active_children()}

    with pytest.raises(TypeError, match="picklable callable or ValidatorSpec"):
        validate_before_deadline(
            unpicklable,
            tmp_path / "artifact.pdf",
            deadline=time.monotonic() + 2,
        )

    after = {child.pid for child in artifact_transport.multiprocessing.active_children()}
    assert after == before


@pytest.mark.skipif(
    os.name != "posix", reason="/dev/fd fd-growth check is POSIX-only"
)
def test_deadline_validator_normalizes_lambda_pickling_error_without_leaks(
    tmp_path, monkeypatch
):
    before_children = {
        child.pid for child in artifact_transport.multiprocessing.active_children()
    }
    before_fds = len(list(Path("/dev/fd").iterdir()))

    def reject_local_callable(_value):
        raise pickle.PicklingError("local lambda is not importable")

    monkeypatch.setattr(
        artifact_transport.ForkingPickler,
        "dumps",
        staticmethod(reject_local_callable),
    )

    with pytest.raises(TypeError, match="picklable callable or ValidatorSpec"):
        validate_before_deadline(
            lambda _path: None,
            tmp_path / "artifact.pdf",
            deadline=time.monotonic() + 2,
        )

    after_children = {
        child.pid for child in artifact_transport.multiprocessing.active_children()
    }
    after_fds = len(list(Path("/dev/fd").iterdir()))
    assert after_children == before_children
    assert after_fds <= before_fds


def test_deadline_validator_closes_pipe_when_process_construction_fails(
    tmp_path, monkeypatch
):
    class Endpoint:
        closed = False

        def close(self):
            self.closed = True

    receiver = Endpoint()
    sender = Endpoint()

    class FailedContext:
        def Pipe(self, *, duplex):
            assert duplex is False
            return receiver, sender

        def Process(self, **_kwargs):
            raise OSError("constructor failed")

    monkeypatch.setattr(
        artifact_transport.multiprocessing,
        "get_context",
        lambda method: FailedContext(),
    )

    with pytest.raises(OSError, match="constructor failed"):
        validate_before_deadline(
            _accept_validator,
            tmp_path / "artifact.pdf",
            deadline=time.monotonic() + 2,
        )

    assert receiver.closed
    assert sender.closed


def test_deadline_validator_closes_pipe_when_spawn_start_fails(
    tmp_path, monkeypatch
):
    class Endpoint:
        closed = False

        def close(self):
            self.closed = True

    receiver = Endpoint()
    sender = Endpoint()

    class FailedProcess:
        def start(self):
            raise OSError("spawn unavailable")

        def is_alive(self):
            return False

    class FailedContext:
        def Pipe(self, *, duplex):
            assert duplex is False
            return receiver, sender

        def Process(self, **_kwargs):
            return FailedProcess()

    monkeypatch.setattr(
        artifact_transport.multiprocessing,
        "get_context",
        lambda method: FailedContext(),
    )

    with pytest.raises(OSError, match="spawn unavailable"):
        validate_before_deadline(
            _accept_validator,
            tmp_path / "artifact.pdf",
            deadline=time.monotonic() + 2,
        )

    assert receiver.closed
    assert sender.closed


def test_copy_file_preserves_preexisting_target_on_exclusive_create_failure(tmp_path):
    source = tmp_path / "source.bin"
    target = tmp_path / "target.bin"
    source.write_bytes(b"new")
    target.write_bytes(b"approved-existing")

    with pytest.raises(FileExistsError):
        copy_file_before_deadline(
            source, target, deadline=time.monotonic() + 1
        )

    assert target.read_bytes() == b"approved-existing"


def test_copy_file_rejects_identical_source_and_target_without_deleting_it(tmp_path):
    source = tmp_path / "source.bin"
    source.write_bytes(b"approved-existing")

    with pytest.raises(ValueError, match="different paths"):
        copy_file_before_deadline(
            source, source, deadline=time.monotonic() + 1
        )

    assert source.read_bytes() == b"approved-existing"
