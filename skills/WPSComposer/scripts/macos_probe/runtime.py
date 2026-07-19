from __future__ import annotations

import json
import os
import plistlib
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path
from typing import List, Optional

WPS_APP = Path("/Applications/wpsoffice.app")
PUBLISH_XML = Path.home() / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
COMPONENT_CONFIG = {
    "writer": {"addon_type": "wps", "port": 3895, "script": "writer.js"},
    "presentation": {"addon_type": "wpp", "port": 3896, "script": "presentation.js"},
    "spreadsheet": {"addon_type": "et", "port": 3897, "script": "spreadsheet.js"},
}

_SHARED_ASSETS = ("index.html", "manifest.xml", "ribbon.xml", "bridge-client.js")


def build_profile(
    assets_dir: Path,
    profiles_dir: Path,
    component: str,
    bridge_url: str,
    token: str,
) -> Path:
    config = COMPONENT_CONFIG[component]
    profile = profiles_dir / component
    profile.mkdir(parents=True, exist_ok=True)

    for name in _SHARED_ASSETS:
        shutil.copy2(assets_dir / name, profile / name)
    shutil.copy2(assets_dir / config["script"], profile / "component.js")

    package = {
        "name": f"wpscomposer-phase0-{component}",
        "version": "0.1.0",
        "private": True,
        "addonType": config["addon_type"],
    }
    (profile / "package.json").write_text(
        json.dumps(package, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )

    session = {"bridgeUrl": bridge_url, "component": component, "token": token}
    (profile / "session.json").write_text(
        json.dumps(session, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    return profile


class RegistrationSnapshot:
    def __init__(self, path: Path, recovery_dir: Path, original: Optional[bytes]):
        self._path = path
        self._recovery_dir = recovery_dir
        self._original = original

    @classmethod
    def capture(cls, path: Path, recovery_dir: Path) -> "RegistrationSnapshot":
        original = path.read_bytes() if path.is_file() else None
        recovery_dir.mkdir(parents=True, mode=0o700, exist_ok=True)
        meta = {"path": str(path), "existed": original is not None}
        (recovery_dir / "registration.json").write_text(
            json.dumps(meta, indent=2) + "\n", encoding="utf-8"
        )
        if original is not None:
            backup = recovery_dir / "publish.xml.original"
            backup.write_bytes(original)
            backup.chmod(0o600)
        print(f"Recovery directory: {recovery_dir.resolve()}", file=sys.stderr)
        return cls(path, recovery_dir, original)

    def restore(self) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        if self._original is not None:
            tmp = self._path.parent / f".{self._path.name}.tmp"
            with tmp.open("wb") as fh:
                fh.write(self._original)
                fh.flush()
                os.fsync(fh.fileno())
            os.replace(tmp, self._path)
            assert self._path.read_bytes() == self._original
        else:
            self._path.unlink(missing_ok=True)
            assert not self._path.exists()
        for name in ("registration.json", "publish.xml.original"):
            (self._recovery_dir / name).unlink(missing_ok=True)
        self._recovery_dir.rmdir()


def find_node(override: Optional[str] = None) -> Path:
    candidates = []
    if override:
        candidates.append(Path(override))
    env = os.environ.get("WPSCOMPOSER_NODE")
    if env:
        candidates.append(Path(env))
    which = shutil.which("node")
    if which:
        candidates.append(Path(which))

    for candidate in candidates:
        resolved = candidate.resolve()
        if not resolved.is_file():
            continue
        try:
            out = subprocess.run(
                [str(resolved), "--version"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            version = out.stdout.strip().lstrip("v")
            major = int(version.split(".")[0])
            if major >= 20:
                return resolved
        except (OSError, ValueError, subprocess.TimeoutExpired):
            continue
    raise RuntimeError("Node.js 20 or newer is required")


def find_wpsjs_cli(probe_root: Path) -> Path:
    cli = probe_root / "node_modules" / "wpsjs" / "src" / "index.js"
    if not cli.is_file():
        raise RuntimeError(
            f"wpsjs CLI not found at {cli}. Run: cd {probe_root} && npm ci"
        )
    return cli


def read_wps_version() -> str:
    with (WPS_APP / "Contents/Info.plist").open("rb") as stream:
        value = plistlib.load(stream)["CFBundleShortVersionString"]
    return str(value)


def _port_free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def _wait_http(url: str, timeout: float = 15.0) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


class ProbeRuntime:
    def __init__(self, probe_root: Path, run_dir: Path, node: Optional[str] = None):
        if sys.platform != "darwin":
            raise RuntimeError("ProbeRuntime requires macOS")
        if not WPS_APP.is_dir():
            raise RuntimeError(f"WPS not found at {WPS_APP}")
        self._probe_root = probe_root
        self._run_dir = run_dir
        self._node = find_node(node)
        self._wpsjs_cli = find_wpsjs_cli(probe_root)
        self._assets = probe_root / "addin"
        self._profiles_dir = run_dir / "profiles"
        self._logs_dir = run_dir / "logs"
        self._processes: List[subprocess.Popen] = []
        self._snapshot: Optional[RegistrationSnapshot] = None

    def prepare_profiles(self, bridge_url: str, token: str) -> None:
        self._run_dir.mkdir(parents=True, exist_ok=True)
        self._logs_dir.mkdir(parents=True, exist_ok=True)
        recovery = self._run_dir / "recovery"
        self._snapshot = RegistrationSnapshot.capture(PUBLISH_XML, recovery)
        for component in ("writer", "presentation", "spreadsheet"):
            build_profile(self._assets, self._profiles_dir, component, bridge_url, token)

    def start_servers(self) -> None:
        self._actual_ports: dict[str, int] = {}
        for component in ("writer", "presentation", "spreadsheet"):
            config = COMPONENT_CONFIG[component]
            profile = self._profiles_dir / component
            log_path = self._logs_dir / f"{component}.log"
            log_file = log_path.open("w")
            proc = subprocess.Popen(
                [
                    str(self._node),
                    str(self._wpsjs_cli),
                    "debug",
                    "--server",
                ],
                cwd=str(profile),
                stdout=log_file,
                stderr=subprocess.STDOUT,
            )
            self._processes.append(proc)
            port = self._detect_port(log_path, timeout=15)
            if port is None:
                raise RuntimeError(f"wpsjs server for {component} did not start")
            self._actual_ports[component] = port

    def _detect_port(self, log_path: Path, timeout: float = 15.0) -> Optional[int]:
        import re
        deadline = time.monotonic() + timeout
        pattern = re.compile(r"http://127\.0\.0\.1:(\d+)")
        while time.monotonic() < deadline:
            try:
                text = log_path.read_text(encoding="utf-8", errors="replace")
                match = pattern.search(text)
                if match:
                    port = int(match.group(1))
                    if _wait_http(f"http://127.0.0.1:{port}/index.html", timeout=5):
                        return port
            except OSError:
                pass
            time.sleep(0.5)
        return None

    def activate_components(self) -> None:
        res_dir = self._probe_root / "node_modules" / "wpsjs" / "src" / "lib" / "res"
        fixtures = {"wpsDemo.docx": "writer", "wppDemo.pptx": "presentation", "etDemo.xlsx": "spreadsheet"}
        for fixture_name in fixtures:
            src = res_dir / fixture_name
            if src.is_file():
                dst = self._run_dir / fixture_name
                shutil.copy2(src, dst)
                subprocess.run(["open", "-a", str(WPS_APP), str(dst)], check=False)
                time.sleep(2)

    def restore_registration(self) -> None:
        for proc in self._processes:
            proc.terminate()
        deadline = time.monotonic() + 5
        for proc in self._processes:
            remaining = max(0.1, deadline - time.monotonic())
            try:
                proc.wait(timeout=remaining)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
        self._processes.clear()
        if self._snapshot:
            self._snapshot.restore()
            self._snapshot = None
