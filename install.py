#!/usr/bin/env python3
"""Install WPSComposer into a personal Codex plugin marketplace."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import tempfile
from typing import Optional


PLUGIN_NAME = "wps-composer"
IGNORED_NAMES = {
    ".git",
    ".pytest_cache",
    ".venv",
    "__pycache__",
    "build",
    "dist",
    "docs",
    "node_modules",
    "tests",
}


class InstallerError(RuntimeError):
    """Raised when the local plugin installation cannot be completed safely."""


@dataclass(frozen=True)
class InstallResult:
    destination: Path
    marketplace_file: Path
    source_path: str


def _validate_source(source_root: Path) -> None:
    required = [
        source_root / ".codex-plugin" / "plugin.json",
        source_root / "skills" / "WPSComposer" / "SKILL.md",
    ]
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise InstallerError(
            "invalid WPSComposer source bundle; missing: " + ", ".join(missing)
        )


def _read_marketplace(path: Path) -> dict:
    if not path.exists():
        return {
            "name": "personal",
            "interface": {"displayName": "Personal"},
            "plugins": [],
        }
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise InstallerError(f"invalid marketplace JSON at {path}: {exc}") from exc
    if not isinstance(data, dict) or not isinstance(data.get("plugins"), list):
        raise InstallerError(
            f"invalid marketplace JSON at {path}: 'plugins' must be a list"
        )
    return data


def _marketplace_source_path(destination: Path, marketplace_root: Path) -> str:
    try:
        relative = destination.resolve().relative_to(marketplace_root.resolve())
    except ValueError as exc:
        raise InstallerError(
            "plugin destination must be inside the marketplace root"
        ) from exc
    return "./" + relative.as_posix()


def _plugin_entry(source_path: str) -> dict:
    return {
        "name": PLUGIN_NAME,
        "source": {"source": "local", "path": source_path},
        "policy": {
            "installation": "AVAILABLE",
            "authentication": "ON_INSTALL",
        },
        "category": "Productivity",
    }


def _write_json_atomically(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=path.name + ".", suffix=".tmp", dir=path.parent
    )
    temporary_path = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary_path, path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _restore_bytes_atomically(path: Path, original: Optional[bytes]) -> None:
    """Restore exact pre-transaction bytes without relying on JSON writing."""
    if original is None:
        path.unlink(missing_ok=True)
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=path.name + ".restore.", suffix=".tmp", dir=path.parent
    )
    temporary_path = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as handle:
            handle.write(original)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_path, path)
    finally:
        temporary_path.unlink(missing_ok=True)


def _copy_plugin(source_root: Path, destination: Path, force: bool) -> None:
    if destination.exists():
        if not force:
            raise InstallerError(
                f"plugin destination already exists: {destination}; use --force"
            )
        shutil.rmtree(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(
        source_root,
        destination,
        ignore=shutil.ignore_patterns(*sorted(IGNORED_NAMES), "*.pyc", "*.pyo"),
    )


def _install_macos_runtime(destination: Path) -> None:
    """Install the pinned production JSAPI runtime into a staged plugin."""
    if platform.system() != "Darwin":
        return
    probe_root = destination / "macos" / "wps-jsapi-probe"
    if not (probe_root / "package.json").is_file():
        return
    raw_node = shutil.which("node")
    raw_npm = shutil.which("npm")
    if not raw_node or not raw_npm:
        raise InstallerError(
            "Node.js 20+ and npm are required for macOS WPS conversion"
        )
    node = Path(raw_node).expanduser().absolute()
    sibling_npm = node.parent / "npm"
    npm = (
        sibling_npm
        if sibling_npm.is_file()
        else Path(raw_npm).expanduser().absolute()
    )
    try:
        version = subprocess.run(
            [str(node), "--version"],
            check=True,
            capture_output=True,
            text=True,
            timeout=10,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        raise InstallerError(f"failed to validate Node.js: {exc}") from exc
    match = re.match(r"^v(\d+)", version.stdout.strip())
    if not match or int(match.group(1)) < 20:
        raise InstallerError(
            f"Node.js 20 or newer is required; found {version.stdout.strip()!r}"
        )
    try:
        subprocess.run(
            [str(npm), "ci", "--omit=dev", "--ignore-scripts"],
            cwd=probe_root,
            check=True,
            capture_output=True,
            text=True,
            timeout=180,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        details = getattr(exc, "stderr", "") or str(exc)
        raise InstallerError(f"failed to install macOS WPS runtime: {details}") from exc
    _write_json_atomically(probe_root / "runtime.json", {"node": str(node)})


def install_plugin(
    source_root: Path,
    codex_home: Path,
    marketplace_root: Path,
    force: bool = False,
) -> InstallResult:
    source_root = source_root.resolve()
    codex_home = codex_home.expanduser().resolve()
    marketplace_root = marketplace_root.expanduser().resolve()
    _validate_source(source_root)

    destination = codex_home / "plugins" / PLUGIN_NAME
    marketplace_file = (
        marketplace_root / ".agents" / "plugins" / "marketplace.json"
    )
    marketplace = _read_marketplace(marketplace_file)
    marketplace_original = (
        marketplace_file.read_bytes() if marketplace_file.exists() else None
    )
    source_path = _marketplace_source_path(destination, marketplace_root)

    if destination.exists() and not force:
        raise InstallerError(
            f"plugin destination already exists: {destination}; use --force"
        )

    destination.parent.mkdir(parents=True, exist_ok=True)
    staging_root = Path(
        tempfile.mkdtemp(prefix=".wpscomposer-install-", dir=destination.parent)
    )
    staged_destination = staging_root / PLUGIN_NAME
    previous_destination = staging_root / "previous-plugin"
    previous_moved = False
    replacement_installed = False
    try:
        _copy_plugin(source_root, staged_destination, force=False)
        _install_macos_runtime(staged_destination)
        if destination.exists():
            os.replace(destination, previous_destination)
            previous_moved = True
        os.replace(staged_destination, destination)
        replacement_installed = True
        plugins = [
            plugin
            for plugin in marketplace["plugins"]
            if not isinstance(plugin, dict) or plugin.get("name") != PLUGIN_NAME
        ]
        plugins.append(_plugin_entry(source_path))
        marketplace["plugins"] = plugins
        _write_json_atomically(marketplace_file, marketplace)
    except Exception:
        if replacement_installed and destination.exists():
            shutil.rmtree(destination)
        if previous_moved and previous_destination.exists():
            os.replace(previous_destination, destination)
        _restore_bytes_atomically(marketplace_file, marketplace_original)
        raise
    finally:
        shutil.rmtree(staging_root, ignore_errors=True)
    return InstallResult(destination, marketplace_file, source_path)


def _parser() -> argparse.ArgumentParser:
    home = Path.home()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true")
    parser.add_argument(
        "--codex-home",
        type=Path,
        default=Path(os.environ.get("CODEX_HOME", home / ".codex")),
    )
    parser.add_argument("--marketplace-root", type=Path, default=home)
    return parser


def main() -> int:
    args = _parser().parse_args()
    try:
        result = install_plugin(
            Path(__file__).resolve().parent,
            args.codex_home,
            args.marketplace_root,
            force=args.force,
        )
    except InstallerError as exc:
        print(f"error: {exc}")
        return 1
    print(f"Installed plugin: {result.destination}")
    print(f"Updated marketplace: {result.marketplace_file}")
    print(
        "Restart the ChatGPT/Codex desktop app, then install or enable "
        "wps-composer from Plugins."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
