from __future__ import annotations

from pathlib import Path
import shutil
import stat

import pytest

from skills.WPSComposer.scripts.artifact_transport import validate_office_package
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
