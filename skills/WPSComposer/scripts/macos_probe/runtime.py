from __future__ import annotations

from contextlib import contextmanager
from functools import partial
import hashlib
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import json
import os
import platform
import plistlib
import re
import shutil
import signal
import socket
import stat
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from threading import Thread
from typing import BinaryIO, Optional
from urllib.error import URLError
from urllib.request import urlopen

WPS_APP = Path("/Applications/wpsoffice.app")
WPS_STAGING_ROOT = (
    Path.home()
    / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer"
)
PUBLISH_XML = (
    Path.home()
    / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
)
COMPONENT_CONFIG = {
    "writer": {"addon_type": "wps", "port": 3889, "script": "writer.js"},
    "presentation": {
        "addon_type": "wpp",
        "port": 3890,
        "script": "presentation.js",
    },
    "spreadsheet": {
        "addon_type": "et",
        "port": 3891,
        "script": "spreadsheet.js",
    },
}
FIXTURE_NAMES = {
    "writer": "wpsDemo.docx",
    "presentation": "wppDemo.pptx",
    "spreadsheet": "etDemo.xlsx",
}
RUNTIME_LOCK_TIMEOUT = 120.0
SERVER_STARTUP_TIMEOUT = 15.0
ACTIVATION_TIMEOUT = 15.0
CLEANUP_GRACE_SECONDS = 5.0


def remaining(deadline: float) -> float:
    """Return the non-negative budget left on one absolute monotonic deadline."""
    return max(0.0, float(deadline) - time.monotonic())


def require_remaining(deadline: float, message: str = "WPS deadline expired") -> float:
    """Return remaining budget or fail before starting another blocking stage."""
    budget = remaining(deadline)
    if budget <= 0:
        raise TimeoutError(message)
    return budget


@dataclass(frozen=True)
class ProcessIdentity:
    """Stable identity used to revalidate one WPS process before signaling."""

    pid: int
    start_time: str
    executable: str


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
    session_nonce: str,
) -> Path:
    """Create one static add-in profile for the selected WPS component.

    The bridge token is deliberately NOT written here: this directory is
    statically served by ``wpsjs debug --server``, so the add-in claims the
    token from the bridge's ``/v1/session`` endpoint at startup.
    """
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
            "sessionNonce": session_nonce,
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
    managed_names: tuple[str, ...] = ()
    managed_digest: Optional[str] = None

    @classmethod
    def capture(
        cls, path: Path, recovery_dir: Path
    ) -> "RegistrationSnapshot":
        target = path.expanduser().resolve()
        recovery = recovery_dir.expanduser().resolve()
        if recovery.exists():
            cls._recover_stale(recovery)
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
                "managedNames": [],
                "managedDigest": None,
            },
        )
        os.chmod(recovery / "registration.json", 0o600)
        if original is not None:
            backup = recovery / "publish.xml.original"
            backup.write_bytes(original)
            os.chmod(backup, 0o600)
        return cls(target, original, original_mode, recovery)

    def record_managed_state(
        self, names: tuple[str, ...], content: bytes
    ) -> None:
        """Persist the exact registration state owned by this session."""
        self.managed_names = names
        self.managed_digest = hashlib.sha256(content).hexdigest()
        _write_json(
            self.recovery_dir / "registration.json",
            {
                "path": str(self.path),
                "existed": self.original is not None,
                "mode": self.original_mode,
                "managedNames": list(names),
                "managedDigest": self.managed_digest,
            },
        )
        os.chmod(self.recovery_dir / "registration.json", 0o600)

    def restore(self) -> None:
        for _ in range(5):
            current = self.path.read_bytes() if self.path.is_file() else None
            current_digest = (
                hashlib.sha256(current).hexdigest() if current is not None else None
            )
            if self.managed_digest is not None and current_digest == self.managed_digest:
                restored = self.original
            elif current is None:
                # An external actor removed the file. Preserve that edit.
                restored = None
            else:
                restored = self._without_managed_entries(current)
            if restored is None:
                if current is None:
                    self._remove_recovery_files()
                    return
                if not self.path.is_file() or self.path.read_bytes() != current:
                    continue
                self.path.unlink()
                self._remove_recovery_files()
                return
            if _atomic_write(
                self.path,
                restored,
                self.original_mode,
                expected=current,
            ):
                self._remove_recovery_files()
                return
        raise RuntimeError(
            f"Registration changed repeatedly; recovery retained at {self.recovery_dir}"
        )

    def _without_managed_entries(self, current: bytes) -> bytes:
        if not self.managed_names:
            return current
        try:
            root = ET.fromstring(current)
        except ET.ParseError as exc:
            raise RuntimeError(
                f"Registration changed to invalid XML; recovery retained at "
                f"{self.recovery_dir}"
            ) from exc
        managed = set(self.managed_names)
        for child in list(root):
            if (
                child.tag.rsplit("}", 1)[-1] == "jspluginonline"
                and child.attrib.get("name") in managed
            ):
                root.remove(child)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    def _replace_target(self, content: Optional[bytes]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if content is None:
            self.path.unlink(missing_ok=True)
            return
        temp_name: Optional[str] = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                dir=self.path.parent,
                prefix=".wpscomposer-restore-",
                delete=False,
            ) as stream:
                stream.write(content)
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
        if self.path.read_bytes() != content:
            raise RuntimeError(f"Registration restore verification failed: {self.path}")

    def _remove_recovery_files(self) -> None:
        for name in ("publish.xml.original", "registration.json"):
            (self.recovery_dir / name).unlink(missing_ok=True)
        self.recovery_dir.rmdir()

    @classmethod
    def _recover_stale(cls, recovery: Path) -> None:
        """Restore registration left behind by a crashed probe run."""
        meta_path = recovery / "registration.json"
        if not meta_path.is_file():
            raise RuntimeError(
                f"Incomplete recovery directory at {recovery}; "
                "inspect and remove it manually."
            )
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        target = Path(meta["path"])
        backup = recovery / "publish.xml.original"
        existed = bool(meta.get("existed"))
        if existed and not backup.is_file():
            # crash between writing metadata and the backup: restoring would
            # DELETE the user's real publish.xml — refuse
            raise RuntimeError(
                f"Recovery metadata at {recovery} is missing its backup; "
                "inspect and remove it manually."
            )
        original = backup.read_bytes() if existed else None
        if (
            original is not None
            and target.is_file()
            and target.read_bytes() == original
        ):
            # crash before registration was ever modified: nothing to restore
            meta_path.unlink(missing_ok=True)
            backup.unlink(missing_ok=True)
            recovery.rmdir()
            return
        snapshot = cls(
            target,
            original,
            meta.get("mode"),
            recovery,
            tuple(meta.get("managedNames", ())),
            meta.get("managedDigest"),
        )
        snapshot.restore()


_UNCONDITIONAL = object()


def _atomic_write(
    path: Path,
    content: bytes,
    mode: Optional[int],
    *,
    expected: object = _UNCONDITIONAL,
) -> bool:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_name: Optional[str] = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", dir=path.parent, prefix=".wpscomposer-write-", delete=False
        ) as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
            temp_name = stream.name
        os.chmod(temp_name, mode if mode is not None else 0o600)
        if expected is not _UNCONDITIONAL:
            current = path.read_bytes() if path.is_file() else None
            if current != expected:
                return False
        os.replace(temp_name, path)
        temp_name = None
        return True
    finally:
        if temp_name is not None:
            Path(temp_name).unlink(missing_ok=True)


def install_registration_entries(
    snapshot: RegistrationSnapshot,
    component_config: dict[str, dict[str, object]],
    *,
    session_nonce: str,
) -> None:
    """Merge this session's uniquely named add-ins into publish.xml."""
    names = tuple(
        f"wpscomposer-{session_nonce}-{component}"
        for component in component_config
    )
    for _ in range(5):
        existed = snapshot.path.is_file()
        current = snapshot.path.read_bytes() if existed else None
        source = current if current is not None else b"<jsplugins/>"
        try:
            root = ET.fromstring(source)
        except ET.ParseError as exc:
            raise RuntimeError(f"Invalid WPS registration XML: {snapshot.path}") from exc
        # Persist ownership before the global file is changed so crash recovery
        # can identify our entries even if the process dies during publication.
        snapshot.record_managed_state(names, source)
        for name, (component, config) in zip(names, component_config.items()):
            ET.SubElement(
                root,
                "jspluginonline",
                {
                    "name": name,
                    "type": str(config["addon_type"]),
                    "url": f"http://127.0.0.1:{config['port']}/",
                    "debug": "",
                    "enable": "enable_dev",
                    "install": "null",
                },
            )
        content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        if _atomic_write(
            snapshot.path,
            content,
            snapshot.original_mode,
            expected=current,
        ):
            snapshot.record_managed_state(names, content)
            return
    raise RuntimeError(
        f"WPS registration changed repeatedly: {snapshot.path}"
    )


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


def read_configured_node(probe_root: Path) -> Optional[str]:
    """Read the Node executable recorded by the transactional installer."""
    runtime_file = probe_root.resolve() / "runtime.json"
    if not runtime_file.is_file():
        return None
    try:
        payload = json.loads(runtime_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Invalid installed runtime configuration: {runtime_file}") from exc
    raw_node = payload.get("node") if isinstance(payload, dict) else None
    if not isinstance(raw_node, str) or not raw_node:
        raise RuntimeError(f"Invalid installed runtime configuration: {runtime_file}")
    return str(Path(raw_node).expanduser().resolve())


@contextmanager
def wps_runtime_lock(
    lock_path: Path,
    timeout: float = RUNTIME_LOCK_TIMEOUT,
    *,
    deadline: Optional[float] = None,
):
    """Serialize WPS registration and fixed-port ownership across processes."""
    import fcntl

    target = lock_path.expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    descriptor = os.open(target, os.O_CREAT | os.O_RDWR, 0o600)
    if deadline is None:
        deadline = time.monotonic() + timeout
    acquired = False
    try:
        while True:
            try:
                fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
                acquired = True
                break
            except BlockingIOError:
                if remaining(deadline) <= 0:
                    raise TimeoutError(
                        "Timed out waiting for another WPSComposer session"
                    )
                time.sleep(min(0.05, remaining(deadline)))
        yield
    finally:
        if acquired:
            fcntl.flock(descriptor, fcntl.LOCK_UN)
        os.close(descriptor)


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


def activation_command(
    app_path: Path, fixture: Path, *, reuse_running: bool = True
) -> list[str]:
    """Open the fixture document in WPS.

    With ``reuse_running=True`` (default) and WPS already running, the fixture
    is handed to the existing instance (opens as a new tab — no second WPS
    app window, no second Dock icon). Falls back to an isolated ``open -n``
    instance when WPS is not running or reuse is disabled, so a live WPS
    session is never disturbed.
    """
    if reuse_running and list_wps_pids(app_path):
        return ["open", "-a", str(app_path), str(fixture)]
    return ["open", "-n", "-a", str(app_path), str(fixture)]


def list_wps_pids(app_path: Path) -> set[int]:
    """Return PIDs whose command is exactly the selected WPS main executable."""
    return set(list_wps_processes(app_path))


def _parse_process_line(
    raw_line: str, expected_executable: str
) -> Optional[ProcessIdentity]:
    fields = raw_line.strip().split(maxsplit=6)
    if len(fields) != 7 or fields[6] != expected_executable:
        return None
    try:
        pid = int(fields[0])
    except ValueError:
        return None
    return ProcessIdentity(
        pid=pid,
        start_time=" ".join(fields[1:6]),
        executable=fields[6],
    )


def list_wps_processes(
    app_path: Path, timeout: float = CLEANUP_GRACE_SECONDS
) -> dict[int, ProcessIdentity]:
    """Return verifiable identities for the selected WPS main executable."""
    executable = str((app_path / "Contents/MacOS/wpsoffice").resolve())
    result = subprocess.run(
        ["ps", "-axo", "pid=,lstart=,command="],
        check=True,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    processes: dict[int, ProcessIdentity] = {}
    for raw_line in result.stdout.splitlines():
        identity = _parse_process_line(raw_line, executable)
        if identity is not None:
            processes[identity.pid] = identity
    return processes


def read_wps_process_identity(
    pid: int,
    app_path: Path,
    timeout: float = CLEANUP_GRACE_SECONDS,
) -> Optional[ProcessIdentity]:
    """Re-read one PID and return it only if it is still the selected WPS."""
    executable = str((app_path / "Contents/MacOS/wpsoffice").resolve())
    try:
        result = subprocess.run(
            ["ps", "-p", str(pid), "-o", "pid=,lstart=,command="],
            check=True,
            capture_output=True,
            text=True,
            timeout=max(0.001, timeout),
        )
    except (OSError, subprocess.SubprocessError):
        return None
    for raw_line in result.stdout.splitlines():
        identity = _parse_process_line(raw_line, executable)
        if identity is not None and identity.pid == pid:
            return identity
    return None


def owned_wps_pids(before: set[int], after: set[int]) -> set[int]:
    """Identify WPS instances that appeared during this probe run."""
    return after - before


def create_staging_session(root: Path = WPS_STAGING_ROOT) -> Path:
    """Create one private session inside the WPS application container."""
    parent = root.expanduser().resolve()
    container = Path.home() / "Library/Containers/com.kingsoft.wpsoffice.mac"
    if parent.is_relative_to(container.resolve()) and not container.exists():
        raise RuntimeError(
            f"WPS sandbox container is missing: {container}. "
            "Refusing to fabricate it; is WPS Office for Mac installed?"
        )
    parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    os.chmod(parent, 0o700)
    session = Path(tempfile.mkdtemp(prefix="session-", dir=parent))
    os.chmod(session, 0o700)
    return session


class _QuietStaticHandler(SimpleHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        return


class StaticProfileServer:
    """Side-effect-free static server for one add-in profile."""

    def __init__(self, profile: Path, port: int):
        handler = partial(_QuietStaticHandler, directory=str(profile))
        self._server = ThreadingHTTPServer(("127.0.0.1", port), handler)
        self._server.daemon_threads = True
        self._thread = Thread(
            target=self._server.serve_forever,
            name=f"wpscomposer-addin-{port}",
            daemon=True,
        )

    def start(self) -> None:
        self._thread.start()

    def close(self, timeout: float = CLEANUP_GRACE_SECONDS) -> None:
        shutdown = Thread(
            target=self._server.shutdown,
            name="wpscomposer-server-shutdown",
            daemon=True,
        )
        shutdown.start()
        shutdown.join(timeout=max(0.0, timeout))
        self._server.server_close()
        self._thread.join(timeout=0)


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
        staging_root: Path = WPS_STAGING_ROOT,
        deadline: Optional[float] = None,
    ):
        self.probe_root = probe_root.resolve()
        self.runtime_dir = runtime_dir.resolve()
        self.bridge_url = bridge_url
        self.token = token
        self.session_nonce = hashlib.sha256(token.encode("utf-8")).hexdigest()
        self.node_override = node_override
        self.publish_xml = publish_xml.expanduser().resolve()
        self.wps_app = wps_app.resolve()
        self.staging_root = staging_root.expanduser().resolve()
        self.state_dir = self.staging_root.parent / ".wpscomposer-runtime"
        self.recovery_dir = self.state_dir / "registration-recovery"
        self.staging_dir: Optional[Path] = None
        self.profiles: dict[str, Path] = {}
        self.fixtures: dict[str, Path] = {}
        self.logs: dict[str, Path] = {}
        self._processes: list[subprocess.Popen] = []
        self._servers: list[StaticProfileServer] = []
        self._log_streams: list[BinaryIO] = []
        self._snapshot: Optional[RegistrationSnapshot] = None
        self._wps_processes_before: Optional[dict[int, ProcessIdentity]] = None
        self._owned_wps_processes: dict[int, ProcessIdentity] = {}
        self._runtime_lock = None
        self.registration_restored = True
        self.deadline = deadline

    def __enter__(self) -> "ProbeRuntime":
        deadline = self.deadline
        if deadline is None:
            deadline = time.monotonic() + RUNTIME_LOCK_TIMEOUT
            self.deadline = deadline
        self._ensure_runtime_state_dir()
        self._runtime_lock = wps_runtime_lock(
            self.state_dir / "runtime.lock", deadline=deadline
        )
        try:
            self._runtime_lock.__enter__()
        except BaseException:
            self._runtime_lock = None
            raise
        try:
            if self.recovery_dir.exists():
                RegistrationSnapshot._recover_stale(self.recovery_dir)
            require_remaining(deadline)
            self._preflight()
            self._wps_processes_before = list_wps_processes(
                self.wps_app, timeout=require_remaining(deadline)
            )
            require_remaining(deadline)
            self.runtime_dir.mkdir(parents=True, exist_ok=False)
            self.staging_dir = create_staging_session(self.staging_root)
            require_remaining(deadline)
        except BaseException:
            try:
                self.close()
            except BaseException:
                # Preserve the original interrupt/exit while close() exhausts
                # its nested finally blocks, including lock release.
                pass
            raise
        return self

    def _ensure_runtime_state_dir(self) -> None:
        self.state_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
        os.chmod(self.state_dir, 0o700)

    def __exit__(self, exc_type, exc, traceback) -> None:
        try:
            self.close()
        except BaseException:
            if exc is None or not self.registration_restored:
                raise
            # Cleanup must not replace the operation's original timeout/error.
            # Registration recovery remains the exception because it needs
            # explicit operator action and Task 5 promises to surface it.

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
                self.session_nonce,
            )
        return dict(self.profiles)

    def start_servers(self, *, deadline: Optional[float] = None) -> None:
        if set(self.profiles) != set(COMPONENT_CONFIG):
            raise RuntimeError("prepare_profiles() must run before start_servers()")
        if deadline is None:
            deadline = self.deadline
        if deadline is None:
            deadline = time.monotonic() + len(COMPONENT_CONFIG) * SERVER_STARTUP_TIMEOUT
        self._ensure_runtime_state_dir()
        require_remaining(deadline)
        recovery = self.recovery_dir
        self._snapshot = RegistrationSnapshot.capture(
            self.publish_xml, recovery
        )
        self.registration_restored = False
        print(
            f"WPS registration recovery: {recovery}",
            file=sys.stderr,
            flush=True,
        )
        try:
            install_registration_entries(
                self._snapshot,
                COMPONENT_CONFIG,
                session_nonce=self.session_nonce,
            )
            for component, config in COMPONENT_CONFIG.items():
                require_remaining(deadline)
                log_path = self.runtime_dir / f"wpsjs-{component}.log"
                self.logs[component] = log_path
                server = StaticProfileServer(
                    self.profiles[component], int(config["port"])
                )
                server.start()
                self._servers.append(server)
                require_remaining(deadline)
                self._wait_for_server(
                    component, int(config["port"]), deadline
                )
        except Exception:
            try:
                self.close()
            except Exception:
                if not self.registration_restored:
                    raise
            raise

    def _wait_for_server(
        self, component: str, port: int, deadline: float
    ) -> None:
        url = f"http://127.0.0.1:{port}/index.html"
        while remaining(deadline) > 0:
            try:
                with urlopen(
                    url, timeout=min(1.0, require_remaining(deadline))
                ) as response:
                    if response.status == 200:
                        return
            except (OSError, URLError):
                budget = remaining(deadline)
                if budget > 0:
                    time.sleep(min(0.1, budget))
        raise TimeoutError(
            f"Timed out waiting for {component} add-in server before deadline; "
            f"see {self.logs[component]}"
        )

    def activate_components(self) -> dict[str, Path]:
        for component in FIXTURE_NAMES:
            self.activate_component(component)
        return dict(self.fixtures)

    def activate_component(
        self, component: str, *, deadline: Optional[float] = None
    ) -> Path:
        if component not in FIXTURE_NAMES:
            raise ValueError(f"Unknown component: {component}")
        if self.staging_dir is None:
            raise RuntimeError("ProbeRuntime must be entered before activation")
        resource_dir = self.probe_root / "node_modules/wpsjs/src/lib/res"
        fixture_dir = self.staging_dir / "fixtures"
        fixture_dir.mkdir(parents=True, exist_ok=True)
        name = FIXTURE_NAMES[component]
        source = resource_dir / name
        if not source.is_file():
            raise RuntimeError(f"Official wpsjs fixture is missing: {source}")
        target = fixture_dir / name
        if not target.is_file():
            shutil.copy2(source, target)
        command = activation_command(self.wps_app, target, reuse_running=False)
        # LaunchServices does not return the spawned WPS PID. Even a single
        # process appearing around `open -n` could be a concurrent user launch,
        # so it is intentionally not claimed or later signaled.
        if deadline is None:
            deadline = self.deadline
        timeout = (
            ACTIVATION_TIMEOUT
            if deadline is None
            else require_remaining(
                deadline, "Timed out before WPS component activation"
            )
        )
        subprocess.run(command, check=True, timeout=timeout)
        if deadline is not None:
            require_remaining(deadline, "Timed out during WPS component activation")
        self.fixtures[component] = target
        return target

    def restore_registration(self) -> None:
        if self._snapshot is not None:
            self._snapshot.restore()
            self._snapshot = None
            self.registration_restored = True

    def _terminate_owned_wps(self, deadline: Optional[float] = None) -> None:
        if not self._owned_wps_processes:
            return
        if deadline is None:
            deadline = time.monotonic() + CLEANUP_GRACE_SECONDS
        owned = tuple(self._owned_wps_processes.values())
        self._owned_wps_processes.clear()
        term_signaled: list[ProcessIdentity] = []
        for identity in sorted(owned, key=lambda item: item.pid):
            budget = remaining(deadline)
            if budget <= 0:
                break
            if read_wps_process_identity(
                identity.pid, self.wps_app, timeout=budget
            ) != identity:
                continue
            try:
                os.kill(identity.pid, signal.SIGTERM)
                term_signaled.append(identity)
            except ProcessLookupError:
                pass
        if not term_signaled:
            return
        remaining_processes = term_signaled
        while time.monotonic() < deadline:
            still_running = []
            for identity in term_signaled:
                budget = remaining(deadline)
                if budget <= 0:
                    break
                if read_wps_process_identity(
                    identity.pid, self.wps_app, timeout=budget
                ) == identity:
                    still_running.append(identity)
            remaining_processes = still_running
            if not remaining_processes:
                return
            budget = remaining(deadline)
            if budget <= 0.1:
                # These identities were revalidated in this iteration. Signal
                # before the cleanup deadline instead of starting another
                # potentially blocking identity lookup after it.
                for identity in remaining_processes:
                    try:
                        os.kill(identity.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                return
            time.sleep(min(0.1, budget))

    def close(self) -> None:
        cleanup_deadline = time.monotonic() + CLEANUP_GRACE_SECONDS
        try:
            for server in reversed(self._servers):
                server.close(timeout=remaining(cleanup_deadline))
            self._servers.clear()
            for process in reversed(self._processes):
                if process.poll() is None:
                    process.terminate()
            for process in reversed(self._processes):
                if process.poll() is not None:
                    continue
                try:
                    process.wait(timeout=remaining(cleanup_deadline))
                except subprocess.TimeoutExpired:
                    process.kill()
                    budget = remaining(cleanup_deadline)
                    if budget > 0:
                        process.wait(timeout=budget)
            self._processes.clear()
            for stream in self._log_streams:
                stream.close()
            self._log_streams.clear()
        finally:
            try:
                self._terminate_owned_wps(cleanup_deadline)
            finally:
                try:
                    self.restore_registration()
                finally:
                    try:
                        if self.staging_dir is not None:
                            staging_dir = self.staging_dir
                            self.staging_dir = None
                            try:
                                shutil.rmtree(staging_dir)
                            except OSError as exc:
                                raise RuntimeError(
                                    "Failed to remove WPS staging session"
                                ) from exc
                    finally:
                        if self._runtime_lock is not None:
                            runtime_lock = self._runtime_lock
                            self._runtime_lock = None
                            runtime_lock.__exit__(None, None, None)
