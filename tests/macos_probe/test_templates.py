from __future__ import annotations

import os
from pathlib import Path
import shutil
import stat

import pytest

from skills.WPSComposer.scripts.artifact_transport import validate_office_package
from skills.WPSComposer.scripts.macos_probe import templates
from skills.WPSComposer.scripts.macos_probe.templates import (
    TemplateError,
    clone_activation_document,
    clone_template,
)


REAL_WRITER_FIXTURE = (
    Path(__file__).parents[2]
    / "macos/wps-jsapi-probe/node_modules/wpsjs/src/lib/res/wpsDemo.docx"
)
NATIVE_WRITER_BLANK = (
    Path(__file__).parents[2]
    / "macos/wps-jsapi-probe/resources/writer-blank.docx"
)


def test_clone_writer_activation_document_uses_pinned_native_blank(tmp_path: Path):
    probe_root = tmp_path / "probe"
    source = probe_root / "resources/writer-blank.docx"
    source.parent.mkdir(parents=True)
    shutil.copy2(NATIVE_WRITER_BLANK, source)
    staging = tmp_path / "session"
    staging.mkdir(mode=0o700)

    cloned = clone_activation_document(probe_root, staging, "writer")

    assert cloned == staging / "wpscomposer-writer-blank.docx"
    assert cloned.read_bytes() == NATIVE_WRITER_BLANK.read_bytes()
    assert stat.S_IMODE(cloned.stat().st_mode) == 0o600
    validate_office_package(cloned, "docx")


def test_clone_writer_activation_document_rejects_changed_seed(tmp_path: Path):
    source = tmp_path / "probe/resources/writer-blank.docx"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"changed seed")
    staging = tmp_path / "session"
    staging.mkdir()

    with pytest.raises(TemplateError, match="digest mismatch"):
        clone_activation_document(tmp_path / "probe", staging, "writer")


def test_activation_clone_removes_partial_publication_when_link_raises(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    probe_root = tmp_path / "probe"
    source = probe_root / "resources/writer-blank.docx"
    source.parent.mkdir(parents=True)
    shutil.copy2(NATIVE_WRITER_BLANK, source)
    staging = tmp_path / "session"
    staging.mkdir()
    real_link = templates.os.link

    def link_then_fail(source_path: Path, target_path: Path) -> None:
        real_link(source_path, target_path)
        raise OSError("publication interrupted")

    monkeypatch.setattr(templates.os, "link", link_then_fail)

    with pytest.raises(TemplateError, match="digest mismatch"):
        clone_activation_document(probe_root, staging, "writer")

    assert list(staging.iterdir()) == []



@pytest.mark.skipif(os.name != "posix", reason="POSIX file-mode assertions")
def test_clone_template_verifies_digest_and_creates_private_copy(tmp_path: Path):
    source = tmp_path / "probe/node_modules/wpsjs/src/lib/res/wpsDemo.docx"
    source.parent.mkdir(parents=True)
    shutil.copy2(REAL_WRITER_FIXTURE, source)
    staging = tmp_path / "session"
    staging.mkdir(mode=0o700)

    cloned = clone_template(tmp_path / "probe", staging, "writer")

    assert cloned == staging / "generated.docx"
    assert cloned.read_bytes() == source.read_bytes()
    assert stat.S_IMODE(cloned.stat().st_mode) == 0o600
    validate_office_package(cloned, "docx")


def test_clone_template_rejects_changed_fixture(tmp_path: Path):
    source = tmp_path / "probe/node_modules/wpsjs/src/lib/res/wpsDemo.docx"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"changed fixture")

    with pytest.raises(TemplateError, match="digest mismatch"):
        clone_template(tmp_path / "probe", tmp_path / "session", "writer")


def test_clone_template_verifies_the_private_copy_when_source_changes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "probe/node_modules/wpsjs/src/lib/res/wpsDemo.docx"
    source.parent.mkdir(parents=True)
    shutil.copy2(REAL_WRITER_FIXTURE, source)
    expected = source.read_bytes()
    staging = tmp_path / "session"
    staging.mkdir(mode=0o700)
    original_validate = templates.validate_office_package

    def mutate_source_after_validation(path: Path, format_name: str) -> None:
        original_validate(path, format_name)
        source.write_bytes(b"replaced after validation")

    monkeypatch.setattr(
        templates, "validate_office_package", mutate_source_after_validation
    )

    cloned = clone_template(tmp_path / "probe", staging, "writer")

    assert cloned.read_bytes() == expected
    validate_office_package(cloned, "docx")
