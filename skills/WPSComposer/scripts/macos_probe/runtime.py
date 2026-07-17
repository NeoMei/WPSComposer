from __future__ import annotations

import json
import os
import platform
import plistlib
import re
import shutil
import socket
import stat
import subprocess
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, Optional
from urllib.error import URLError
from urllib.request import urlopen

WPS_APP = Path("/Applications/wpsoffice.app")
PUBLISH_XML = (
    Path.home()
    / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
)
COMPONENT_CONFIG = {
    "writer": {"addon_type": "wps", "port": 3891, "script": "writer.js"},
    "presentation": {
        "addon_type": "wpp",
        "port": 3892,
        "script": "presentation.js",
    },
    "spreadsheet": {
        "addon_type": "et",
        "port": 3893,
        "script": "spreadsheet.js",
    },
}
FIXTURE_NAMES = {
    "writer": "wpsDemo.docx",
    "presentation": "wppDemo.pptx",
    "spreadsheet": "etDemo.xlsx",
}


def _write_json(path: Path, payload: object) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def build_profile(
    assets: Path,
    profiles_root: Path,
    component: str,
    bridge_url: str,
    token: str,
) -> Path:
    """Create one static add-in profile for the selected WPS component."""
    if component not in COMPONENT_CONFIG:
        raise ValueError(f"Unknown component: {component}")
    config = COMPONENT_CONFIG[component]
    profile = profiles_root / component
    profile.mkdir(parents=True, exist_ok=False)
    for name in ("index.html", "manifest.xml", "ribbon.xml", "bridge-client.js"):
        shutil.copy2(assets / name, profile / name)
    shutil.copy2(assets / str(config["script"]), profile / "component.js")
    _write_json(
        profile / "package.json",
        {
            "name": f"wpscomposer-phase0-{component}",
            "version": "0.1.0",
            "private": True,
            "addonType": config["addon_type"],
        },
    )
    _write_json(
        profile / "session.json",
        {
            "bridgeUrl": bridge_url,
            "component": component,
            "token": token,
        },
    )
    return profile


@dataclass
class RegistrationSnapshot:
    """Private, crash-recoverable snapshot of WPS add-in registration."""

    path: Path
    original: Optional[bytes]
    original_mode: Optional[int]
    recovery_dir: Path

    @classmethod
    def capture(
        cls, path: Path, recovery_dir: Path
    ) -> "RegistrationSnapshot":
        target = path.expanduser().resolve()
        recovery = recovery_dir.expanduser().resolve()
        recovery.mkdir(parents=True, exist_ok=False, mode=0o700)
        os.chmod(recovery, 0o700)
        existed = target.is_file()
        original = target.read_bytes() if existed else None
        original_mode = stat.S_IMODE(target.stat().st_mode) if existed else None
        _write_json(
            recovery / "registration.json",
            {
                "path": str(target),
                "existed": existed,
                "mode": original_mode,
            },
        )
        os.chmod(recovery / "registration.json", 0o600)
        if original is not None:
            backup = recovery / "publish.xml.original"
            backup.write_bytes(original)
            os.chmod(backup, 0o600)
        return cls(target, original, original_mode, recovery)

    def restore(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.original is None:
            self.path.unlink(missing_ok=True)
            if self.path.exists():
                raise RuntimeError(f"Failed to remove probe registration: {self.path}")
        else:
            temp_name: Optional[str] = None
            try:
                with tempfile.NamedTemporaryFile(
                    mode="wb",
                    dir=self.path.parent,
                    prefix=".wpscomposer-restore-",
                    delete=False,
                ) as stream:
                    stream.write(self.original)
                    stream.flush()
                    os.fsync(stream.fileno())
                    temp_name = stream.name
                if self.original_mode is not None:
                    os.chmod(temp_name, self.original_mode)
                os.replace(temp_name, self.path)
                temp_name = None
            finally:
                if temp_name is not None:
                    Path(temp_name).unlink(missing_ok=True)
            if self.path.read_bytes() != self.original:
                raise RuntimeError(f"Registration restore verification failed: {self.path}")
        self._remove_recovery_files()

    def _remove_recovery_files(self) -> None:
        for name in ("publish.xml.original", "registration.json"):
            (self.recovery_dir / name).unlink(missing_ok=True)
        self.recovery_dir.rmdir()


def find_node(override: Optional[str] = None) -> Path:
    """Find a Node.js 20+ executable without changing the system runtime."""
    explicit = override or os.environ.get("WPSCOMPOSER_NODE")
    candidates = [explicit] if explicit else [shutil.which("node")]
    for raw in candidates:
        if not raw:
            continue
        candidate = Path(raw).expanduser().resolve()
        if not candidate.is_file():
            continue
        try:
            result = subprocess.run(
                [str(candidate), "--version"],
                check=True,
                capture_output=True,
                text=True,
                timeout=5,
            )
        except (OSError, subprocess.SubprocessError):
            continue
        match = re.match(r"^v(\d+)", result.stdout.strip())
        if match and int(match.group(1)) >= 20:
            return candidate
    if explicit:
        raise RuntimeError(f"Node.js override is unavailable or too old: {explicit}")
    raise RuntimeError("Node.js 20 or newer is required")


def find_wpsjs_cli(probe_root: Path) -> Path:
    cli = (probe_root / "node_modules/wpsjs/src/index.js").resolve()
    if not cli.is_file():
        raise RuntimeError(
            f"wpsjs is not installed. Run `npm ci` in {probe_root.resolve()}"
        )
    return cli


def read_wps_version(app_path: Path = WPS_APP) -> str:
    with (app_path / "Contents/Info.plist").open("rb") as stream:
        value = plistlib.load(stream)["CFBundleShortVersionString"]
    return str(value)


def _require_free_port(port: int) -> None:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            probe.bind(("127.0.0.1", port))
        except OSError as exc:
            raise RuntimeError(f"Required add-in port is already in use: {port}") from exc


class ProbeRuntime:
    """Owns temporary add-in profiles and the child wpsjs servers."""

    def __init__(
        self,
        probe_root: Path,
        runtime_dir: Path,
        bridge_url: str,
        token: str,
        node_override: Optional[str] = None,
        publish_xml: Path = PUBLISH_XML,
        wps_app: Path = WPS_APP,
    ):
        self.probe_root = probe_root.resolve()
        self.runtime_dir = runtime_dir.resolve()
        self.bridge_url = bridge_url
        self.token = token
        self.node_override = node_override
        self.publish_xml = publish_xml.expanduser().resolve()
        self.wps_app = wps_app.resolve()
        self.profiles: dict[str, Path] = {}
        self.logs: dict[str, Path] = {}
        self._processes: list[subprocess.Popen] = []
        self._log_streams: list[BinaryIO] = []
        self._snapshot: Optional[RegistrationSnapshot] = None

    def __enter__(self) -> "ProbeRuntime":
        self._preflight()
        self.runtime_dir.mkdir(parents=True, exist_ok=False)
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.close()

    def _preflight(self) -> None:
        if platform.system() != "Darwin":
            raise RuntimeError("The WPS JSAPI probe requires macOS")
        if not self.wps_app.is_dir():
            raise RuntimeError(f"WPS Office is unavailable: {self.wps_app}")
        for config in COMPONENT_CONFIG.values():
            _require_free_port(int(config["port"]))

    def prepare_profiles(self) -> dict[str, Path]:
        assets = self.probe_root / "addin"
        profiles_root = self.runtime_dir / "profiles"
        for component in COMPONENT_CONFIG:
            self.profiles[component] = build_profile(
                assets,
                profiles_root,
                component,
                self.bridge_url,
                self.token,
            )
        return dict(self.profiles)

    def start_servers(self) -> None:
        if set(self.profiles) != set(COMPONENT_CONFIG):
            raise RuntimeError("prepare_profiles() must run before start_servers()")
        node = find_node(self.node_override)
        cli = find_wpsjs_cli(self.probe_root)
        recovery = self.runtime_dir / "registration-recovery"
        self._snapshot = RegistrationSnapshot.capture(
            self.publish_xml, recovery
        )
        print(f"WPS registration recovery: {recovery}", flush=True)
        try:
            for component, config in COMPONENT_CONFIG.items():
                log_path = self.runtime_dir / f"wpsjs-{component}.log"
                log_stream = log_path.open("ab")
                self._log_streams.append(log_stream)
                self.logs[component] = log_path
                environment = os.environ.copy()
                environment["PATH"] = (
                    str(node.parent)
                    + os.pathsep
                    + environment.get("PATH", "")
                )
                process = subprocess.Popen(
                    [
                        str(node),
                        str(cli),
                        "debug",
                        "--server",
                        "--port",
                        str(config["port"]),
                    ],
                    cwd=self.profiles[component],
                    env=environment,
                    stdin=subprocess.DEVNULL,
                    stdout=log_stream,
                    stderr=subprocess.STDOUT,
                )
                self._processes.append(process)
                self._wait_for_server(component, int(config["port"]))
        except Exception:
            self.close()
            raise

    def _wait_for_server(self, component: str, port: int) -> None:
        deadline = time.monotonic() + 15
        url = f"http://127.0.0.1:{port}/index.html"
        process = self._processes[-1]
        while time.monotonic() < deadline:
            if process.poll() is not None:
                raise RuntimeError(
                    f"wpsjs {component} exited early; see {self.logs[component]}"
                )
            try:
                with urlopen(url, timeout=1) as response:
                    if response.status == 200:
                        return
            except (OSError, URLError):
                time.sleep(0.1)
        raise TimeoutError(
            f"Timed out waiting for {component} add-in server; "
            f"see {self.logs[component]}"
        )

    def activate_components(self) -> dict[str, Path]:
        resource_dir = self.probe_root / "node_modules/wpsjs/src/lib/res"
        fixture_dir = self.runtime_dir / "fixtures"
        fixture_dir.mkdir(parents=True, exist_ok=False)
        fixtures: dict[str, Path] = {}
        for component, name in FIXTURE_NAMES.items():
            source = resource_dir / name
            if not source.is_file():
                raise RuntimeError(f"Official wpsjs fixture is missing: {source}")
            target = fixture_dir / name
            shutil.copy2(source, target)
            subprocess.run(
                ["open", "-a", str(self.wps_app), str(target)],
                check=True,
                timeout=15,
            )
            fixtures[component] = target
        return fixtures

    def restore_registration(self) -> None:
        if self._snapshot is not None:
            snapshot = self._snapshot
            self._snapshot = None
            snapshot.restore()

    def close(self) -> None:
        try:
            for process in reversed(self._processes):
                if process.poll() is None:
                    process.terminate()
            for process in reversed(self._processes):
                if process.poll() is not None:
                    continue
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=5)
            self._processes.clear()
            for stream in self._log_streams:
                stream.close()
            self._log_streams.clear()
        finally:
            self.restore_registration()
