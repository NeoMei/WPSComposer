from __future__ import annotations

import pytest

from skills.WPSComposer.scripts.longform.pdf_quality import (
    QualityDependencyError,
    require_quality_dependencies,
)


def test_dependency_gate_accepts_required_versions():
    versions = {"Pillow": "10.0.0", "pypdf": "4.0.0", "pdfplumber": "0.11.0"}
    imported = []

    result = require_quality_dependencies(
        importer=lambda name: imported.append(name) or object(),
        version_getter=versions.__getitem__,
    )

    assert result == ("Pillow 10.0.0", "pypdf 4.0.0", "pdfplumber 0.11.0")
    assert imported == ["PIL", "pypdf", "pdfplumber"]


@pytest.mark.parametrize("missing", ["PIL", "pypdf", "pdfplumber"])
def test_dependency_gate_fails_cleanly_before_native_work(missing):
    calls = []

    def importer(name):
        calls.append(name)
        if name == missing:
            raise ImportError(name)
        return object()

    with pytest.raises(QualityDependencyError, match="quality dependencies unavailable") as exc:
        require_quality_dependencies(
            importer=importer,
            version_getter=lambda package: "99.0",
        )

    assert missing in exc.value.missing_modules
    assert "/Users/" not in str(exc.value)
    assert r"C:\\Users\\" not in str(exc.value)


@pytest.mark.parametrize(
    ("package", "version"),
    [("Pillow", "9.5"), ("pypdf", "3.17"), ("pdfplumber", "0.10.4")],
)
def test_dependency_gate_rejects_too_old_versions(package, version):
    versions = {"Pillow": "10.0", "pypdf": "4.0", "pdfplumber": "0.11"}
    versions[package] = version
    with pytest.raises(QualityDependencyError, match="quality dependencies unavailable"):
        require_quality_dependencies(
            importer=lambda name: object(),
            version_getter=versions.__getitem__,
        )
