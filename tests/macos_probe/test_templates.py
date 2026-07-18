from __future__ import annotations

from pathlib import Path
import shutil
import stat

import pytest

from skills.WPSComposer.scripts.artifact_transport import validate_office_package
from skills.WPSComposer.scripts.macos_probe import templates
from skills.WPSComposer.scripts.macos_probe.templates import (
    TemplateError,
    clone_template,
)


REAL_WRITER_FIXTURE = (
    Path(__file__).parents[2]
    / "macos/wps-jsapi-probe/node_modules/wpsjs/src/lib/res/wpsDemo.docx"
)


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
