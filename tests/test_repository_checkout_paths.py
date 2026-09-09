"""Catch tracked names that prevent a normal Windows checkout on any host.

Audit the Git index, not a filesystem walk: untracked native probe output is
not shipped, and staged renames must be tested before committing.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess

import pytest


ROOT = Path(__file__).resolve().parents[1]
EVIDENCE = "docs/verification/microsoft-parity/macos-word-advanced"
ORIGINAL_COMMIT = "34a0658e38c6d5a73a7066df7f49f73b53aceb39"
RESERVED_CHARACTERS = set('<>:"\\|?*')
RESERVED_DEVICES = {"CON", "PRN", "AUX", "NUL"} | {
    prefix + digit for prefix in ("COM", "LPT") for digit in "123456789¹²³"
}


def _git(*args):
    return subprocess.run(
        ["git", *args], cwd=ROOT, check=True, capture_output=True,
    ).stdout


def _tracked_paths():
    return _git("ls-files", "--cached", "-z").decode("utf-8").split("\0")[:-1]


def _windows_checkout_issues(paths):
    # Microsoft naming rules apply to every directory and filename component:
    # https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file
    issues = []
    seen = {}
    for path in sorted(set(paths)):
        parts = path.split("/")
        for index, part in enumerate(parts):
            if not part or part in {".", ".."}:
                issues.append((path, "invalid relative component"))
            if any(char in RESERVED_CHARACTERS or ord(char) < 32 for char in part):
                issues.append((path, "reserved character"))
            if part.endswith((" ", ".")):
                issues.append((path, "trailing space or dot"))
            if part.split(".", 1)[0].rstrip(" ").upper() in RESERVED_DEVICES:
                issues.append((path, "reserved device name"))

            prefix = "/".join(parts[:index + 1])
            kind = "file" if index == len(parts) - 1 else "directory"
            previous = seen.setdefault(prefix.casefold(), (prefix, kind))
            if previous != (prefix, kind):
                issues.append((path, f"case or file/directory collision with {previous[0]}"))
    return issues


def test_git_tracked_paths_allow_windows_checkout():
    paths = _tracked_paths()
    assert paths, "This regression requires a Git checkout with tracked files"
    issues = _windows_checkout_issues(paths)
    assert not issues, "Windows-incompatible tracked paths:\n" + "\n".join(
        f"{path}: {reason}" for path, reason in issues
    )


@pytest.mark.parametrize("component", [
    "resource:image.png", "less<than", "greater>than", 'double"quote',
    "back\\slash", "pipe|name", "question?", "star*", "control\x01", "nul\x00",
    "trailing.", "trailing ", "CON", "nul.txt", "AUX.tar.gz", "prn",
    "COM1.log", "lpt9", "COM¹.txt", "LPT²", "COM³", "..", "",
])
def test_windows_audit_rejects_invalid_components_in_files_and_directories(component):
    assert _windows_checkout_issues([f"evidence/{component}"])
    assert _windows_checkout_issues([f"evidence/{component}/result.json"])


@pytest.mark.parametrize("paths", [
    ["docs/Report.md", "docs/report.md"],
    ["Docs/first.md", "docs/second.md"],
    ["docs/report", "docs/REPORT/page.md"],
    ["docs/report", "docs/report/page.md"],
])
def test_windows_audit_rejects_colliding_files_and_directory_prefixes(paths):
    assert _windows_checkout_issues(paths)


def test_windows_audit_accepts_unicode_spaces_and_nonreserved_similar_names():
    assert not _windows_checkout_issues([
        "docs/技术报告/slide 01.png", ".github/workflows/tests.yml",
        "docs/CONSOLE.md", "docs/COM10.txt", "docs/LPT0.txt", "docs/null.txt",
        "docs/wpsc-rsrc-K9qCW0iohe-tx8yQ.png",
    ])


def test_relocated_evidence_preserves_original_git_bytes():
    manifest_path = ROOT / EVIDENCE / "windows-path-relocations.json"
    assert manifest_path.is_file(), "Historical evidence relocation map is required"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["original_git_commit"] == ORIGINAL_COMMIT
    expected_originals = {
        f"{EVIDENCE}/integrated-{run}/runtime/wpsc-rsrc:{token}.png"
        for run in ("02", "03")
        for token in ("K9qCW0iohe-tx8yQ", "QNP5Ulby7mHPtP88")
    }
    files = manifest["files"]
    assert len(files) == len(expected_originals)
    assert {entry["original_repository_path"] for entry in files} == expected_originals
    stored_paths = [entry["stored_repository_path"] for entry in files]
    assert len(set(stored_paths)) == len(files)
    assert not _windows_checkout_issues(stored_paths)
    tracked = set(_tracked_paths())
    for entry in files:
        original = entry["original_repository_path"]
        stored = entry["stored_repository_path"]
        assert original not in tracked
        assert stored in tracked
        original_bytes = _git("show", f"{ORIGINAL_COMMIT}:{original}")
        stored_bytes = (ROOT / stored).read_bytes()
        assert stored_bytes == original_bytes
        assert hashlib.sha256(stored_bytes).hexdigest() == entry["sha256"]
        assert len(stored_bytes) == entry["size_bytes"]
        assert entry["unchanged_historical_references"]
        for reference in entry["unchanged_historical_references"]:
            reference_bytes = (ROOT / reference).read_bytes()
            assert reference_bytes == _git("show", f"{ORIGINAL_COMMIT}:{reference}")
            assert entry["original_runtime_path"].encode("utf-8") in reference_bytes
