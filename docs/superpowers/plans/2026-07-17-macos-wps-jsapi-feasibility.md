# macOS WPS JSAPI Phase 0 Feasibility Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a reversible macOS feasibility probe that uses the installed WPS JS add-in runtime to produce DOCX, PPTX, XLSX, and standalone PDF smoke artifacts and records exact JSAPI capability gaps.

**Architecture:** Keep Phase 0 isolated from the production `generate()` path. A Python harness starts an authenticated loopback HTTP bridge, creates three temporary WPS add-in profiles (Writer, Presentation, Spreadsheet), invokes the official `wpsjs 2.2.3` macOS debug path, opens the official WPS fixture files to activate each component, sends typed smoke commands by long polling, validates the four artifacts, writes a capability report, and restores the user's original WPS `publish.xml` in `finally` cleanup.

**Tech Stack:** Python 3.9+ standard library, pytest 8+, JavaScript running inside Mac WPS, `wpsjs==2.2.3`, `wps-jsapi-declare==2.2.0`, WPS Office for Mac at `/Applications/wpsoffice.app`.

## Global Constraints

- Phase 0 is a feasibility gate; do not route the public `generate()` API to macOS yet.
- WPS must perform all DOCX, PPTX, XLSX, and PDF serialization; do not add a pure OOXML, Chromium, LibreOffice, or reportlab output fallback.
- Bind the bridge only to `127.0.0.1`, generate a random token for every run, and require `Authorization: Bearer <token>` on every non-preflight protocol endpoint; CORS `OPTIONS` validates the exact origin before returning no content.
- Accept only the typed methods `probe_capabilities`, `smoke_docx`, `smoke_pdf`, `smoke_pptx`, and `smoke_xlsx`; never accept JavaScript source through the bridge.
- Canonicalize every image and output path and reject any path outside the per-run temporary directory or requested output directory.
- Preserve the existing Windows COM code and keep all current public API tests passing.
- Preserve the user's existing WPS add-in registration file byte-for-byte after the probe exits, including error and timeout paths.
- Never close unrelated WPS documents or quit the user's WPS application; close only documents created by the probe.
- Produce four independent requested outputs: DOCX does not imply PDF, PPTX does not imply PDF, and XLSX does not imply PDF.
- Pin the feasibility toolchain to `wpsjs==2.2.3` and `wps-jsapi-declare==2.2.0`; commit `package-lock.json`.
- Use the `obsidian-reading` label for the standalone PDF smoke command and do not report its temporary DOCX as an output.
- Record WPS version `12.1.26035` as the initially tested local version, while reading and reporting the actual installed version at runtime.

## File Map

- `macos/wps-jsapi-probe/package.json` — pinned Node development dependencies and probe scripts.
- `macos/wps-jsapi-probe/package-lock.json` — reproducible npm dependency graph.
- `macos/wps-jsapi-probe/addin/index.html` — shared add-in HTML entrypoint.
- `macos/wps-jsapi-probe/addin/manifest.xml` — minimal WPS add-in manifest.
- `macos/wps-jsapi-probe/addin/ribbon.xml` — add-in load callback with no visible controls.
- `macos/wps-jsapi-probe/addin/bridge-client.js` — authenticated register/long-poll/result protocol.
- `macos/wps-jsapi-probe/addin/writer.js` — Writer capability, DOCX, and PDF smoke handlers.
- `macos/wps-jsapi-probe/addin/presentation.js` — Presentation capability and PPTX smoke handler.
- `macos/wps-jsapi-probe/addin/spreadsheet.js` — Spreadsheet capability and XLSX smoke handler.
- `skills/WPSComposer/scripts/macos_probe/models.py` — typed command/result models, method routing, and path policy.
- `skills/WPSComposer/scripts/macos_probe/bridge.py` — loopback HTTP bridge and in-memory command state.
- `skills/WPSComposer/scripts/macos_probe/runtime.py` — temporary profiles, tool discovery, `wpsjs` processes, WPS activation, and registration rollback.
- `skills/WPSComposer/scripts/macos_probe/runner.py` — Phase 0 orchestration, artifact validation, and report generation.
- `skills/WPSComposer/scripts/macos_probe/__main__.py` — `python3 -m` command entrypoint.
- `tests/macos_probe/test_models.py` — command and path policy tests.
- `tests/macos_probe/test_bridge.py` — bridge authentication, CORS, queue, and timeout tests.
- `tests/macos_probe/test_runtime.py` — profile creation, Node/tool discovery, and rollback tests.
- `tests/macos_probe/test_runner.py` — orchestration and artifact validation tests with fake component clients.
- `tests/macos_probe/test_addin_assets.py` — static contract tests for manifests and allowed JS command handlers.
- `docs/macos-phase0.md` — exact setup, run, recovery, evidence, and result interpretation.
- `README.md` — short macOS experimental-status link.

---

### Task 1: Define the typed probe protocol and path policy

**Files:**
- Modify: `.gitignore`
- Create: `skills/WPSComposer/scripts/macos_probe/__init__.py`
- Create: `skills/WPSComposer/scripts/macos_probe/models.py`
- Create: `tests/macos_probe/test_models.py`

**Interfaces:**
- Consumes: Python standard library only.
- Produces: `COMPONENTS`, `METHOD_COMPONENT`, `ProbeCommand`, `ProbeResult`, `PathPolicy`, and `ProtocolError`.

- [ ] **Step 1: Create the isolated Python development environment**

Append `.venv/` to `.gitignore`, then run:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
```

Expected: `.venv/bin/python -m pytest --version` reports pytest 8 or newer and `.venv` is not shown by `git status --short`.

- [ ] **Step 2: Write the failing model tests**

```python
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.macos_probe.models import (
    METHOD_COMPONENT,
    PathPolicy,
    ProbeCommand,
    ProbeResult,
    ProtocolError,
)


def test_methods_are_routed_to_one_component():
    assert METHOD_COMPONENT == {
        "probe_capabilities": None,
        "smoke_docx": "writer",
        "smoke_pdf": "writer",
        "smoke_pptx": "presentation",
        "smoke_xlsx": "spreadsheet",
    }


def test_command_rejects_unknown_method():
    with pytest.raises(ProtocolError, match="Unsupported method"):
        ProbeCommand.create("writer", "eval", {})


def test_command_rejects_wrong_component():
    with pytest.raises(ProtocolError, match="requires presentation"):
        ProbeCommand.create("writer", "smoke_pptx", {})


def test_result_round_trip():
    result = ProbeResult("cmd-1", True, {"path": "/tmp/a.docx"}, None)
    assert ProbeResult.from_dict(result.to_dict()) == result


def test_path_policy_accepts_only_declared_roots(tmp_path: Path):
    runtime = tmp_path / "runtime"
    output = tmp_path / "output"
    runtime.mkdir()
    output.mkdir()
    policy = PathPolicy((runtime, output))

    assert policy.require_allowed(output / "smoke.docx") == (
        output / "smoke.docx"
    ).resolve()

    with pytest.raises(ProtocolError, match="outside allowed roots"):
        policy.require_allowed(tmp_path.parent / "escape.docx")
```

- [ ] **Step 3: Run the model tests and verify collection fails**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_models.py -q`

Expected: FAIL during collection with `ModuleNotFoundError: No module named 'skills.WPSComposer.scripts.macos_probe'`.

- [ ] **Step 4: Implement the protocol models exactly**

Create `skills/WPSComposer/scripts/macos_probe/__init__.py`:

```python
"""Isolated macOS WPS JSAPI feasibility probe."""

from .models import ProbeCommand, ProbeResult, ProtocolError

__all__ = ["ProbeCommand", "ProbeResult", "ProtocolError"]
```

Create `skills/WPSComposer/scripts/macos_probe/models.py`:

```python
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Optional, Union
from uuid import uuid4

COMPONENTS = ("writer", "presentation", "spreadsheet")
METHOD_COMPONENT = {
    "probe_capabilities": None,
    "smoke_docx": "writer",
    "smoke_pdf": "writer",
    "smoke_pptx": "presentation",
    "smoke_xlsx": "spreadsheet",
}


class ProtocolError(ValueError):
    """Raised when a bridge message violates the probe protocol."""


@dataclass(frozen=True)
class ProbeCommand:
    id: str
    component: str
    method: str
    params: Mapping[str, Any]

    @classmethod
    def create(
        cls, component: str, method: str, params: Mapping[str, Any]
    ) -> "ProbeCommand":
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")
        if method not in METHOD_COMPONENT:
            raise ProtocolError(f"Unsupported method: {method}")
        required = METHOD_COMPONENT[method]
        if required is not None and required != component:
            raise ProtocolError(f"{method} requires {required}, got {component}")
        return cls(uuid4().hex, component, method, dict(params))

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "component": self.component,
            "method": self.method,
            "params": dict(self.params),
        }


@dataclass(frozen=True)
class ProbeResult:
    id: str
    ok: bool
    value: Mapping[str, Any]
    error: Optional[Mapping[str, Any]]

    @classmethod
    def from_dict(cls, raw: Mapping[str, Any]) -> "ProbeResult":
        command_id = str(raw.get("id", ""))
        if not command_id:
            raise ProtocolError("Result id is required")
        ok = raw.get("ok")
        if not isinstance(ok, bool):
            raise ProtocolError("Result ok must be boolean")
        value = raw.get("value") or {}
        error = raw.get("error")
        if not isinstance(value, dict):
            raise ProtocolError("Result value must be an object")
        if error is not None and not isinstance(error, dict):
            raise ProtocolError("Result error must be an object or null")
        return cls(command_id, ok, value, error)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "ok": self.ok,
            "value": dict(self.value),
            "error": None if self.error is None else dict(self.error),
        }


class PathPolicy:
    def __init__(self, roots: tuple[Path, ...]):
        self._roots = tuple(root.resolve() for root in roots)

    def require_allowed(self, value: Union[str, Path]) -> Path:
        candidate = Path(value).expanduser().resolve()
        if not any(candidate == root or root in candidate.parents for root in self._roots):
            raise ProtocolError(f"Path is outside allowed roots: {candidate}")
        return candidate
```

- [ ] **Step 5: Run the model tests**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_models.py -q`

Expected: `5 passed`.

- [ ] **Step 6: Commit the protocol boundary**

```bash
git add .gitignore skills/WPSComposer/scripts/macos_probe tests/macos_probe/test_models.py
git commit -m "Add typed macOS probe protocol"
```

### Task 2: Build the authenticated loopback bridge

**Files:**
- Create: `skills/WPSComposer/scripts/macos_probe/bridge.py`
- Create: `tests/macos_probe/test_bridge.py`

**Interfaces:**
- Consumes: `ProbeCommand`, `ProbeResult`, `ProtocolError` from Task 1.
- Produces: `BridgeState.issue(component, method, params)`, `BridgeState.wait_result(command_id, timeout)`, host-side `BridgeState.cancel(command_id)`, and context-managed `LoopbackBridge` with public `state`, `url`, `token`, and `wait_registered()`.

- [ ] **Step 1: Write failing authentication and queue tests**

```python
import json
from concurrent.futures import ThreadPoolExecutor
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest

from skills.WPSComposer.scripts.macos_probe.bridge import LoopbackBridge
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult, ProtocolError


def request(bridge, method, path, body=None, token=None, origin=None):
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if origin:
        headers["Origin"] = origin
    data = None if body is None else json.dumps(body).encode()
    req = Request(bridge.url + path, data=data, headers=headers, method=method)
    with urlopen(req, timeout=2) as response:
        raw = response.read()
        return response.status, None if not raw else json.loads(raw)


def test_bridge_rejects_missing_token():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(bridge, "POST", "/v1/register", {"component": "writer"})
        assert error.value.code == 401


def test_bridge_rejects_unlisted_origin():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(
                bridge,
                "POST",
                "/v1/register",
                {"component": "writer"},
                bridge.token,
                "http://example.com",
            )
        assert error.value.code == 403


def test_command_round_trip():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        status, _ = request(
            bridge,
            "POST",
            "/v1/register",
            {"component": "writer"},
            bridge.token,
            origin,
        )
        assert status == 204
        command = bridge.issue("writer", "smoke_docx", {"outputPath": "/tmp/a.docx"})
        status, payload = request(
            bridge,
            "GET",
            "/v1/next?component=writer",
            token=bridge.token,
            origin=origin,
        )
        assert status == 200
        assert payload == command.to_dict()

        status, _ = request(
            bridge,
            "POST",
            "/v1/result",
            {"id": command.id, "ok": True, "value": {"saved": True}, "error": None},
            bridge.token,
            origin,
        )
        assert status == 204
        assert bridge.wait_result(command.id, 1).value == {"saved": True}


def test_wait_registered_reports_independent_components():
    origins = {
        "http://127.0.0.1:3891",
        "http://127.0.0.1:3892",
        "http://127.0.0.1:3893",
    }
    with LoopbackBridge(origins) as bridge, ThreadPoolExecutor() as pool:
        for component, port in (("writer", 3891), ("presentation", 3892), ("spreadsheet", 3893)):
            pool.submit(
                request,
                bridge,
                "POST",
                "/v1/register",
                {"component": component},
                bridge.token,
                f"http://127.0.0.1:{port}",
            )
        bridge.wait_registered({"writer", "presentation", "spreadsheet"}, 2)


def test_completion_is_idempotent_and_cancellation_ignores_late_result():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        first = bridge.issue("writer", "smoke_docx", {})
        result = ProbeResult(first.id, True, {"saved": True}, None)
        bridge.state.complete(result)
        bridge.state.complete(result)
        assert bridge.wait_result(first.id, 1) == result

        second = bridge.issue("writer", "smoke_pdf", {})
        bridge.state.cancel(second.id)
        bridge.state.complete(ProbeResult(second.id, True, {"saved": True}, None))
        with pytest.raises(TimeoutError):
            bridge.wait_result(second.id, 0.01)


def test_conflicting_duplicate_result_is_rejected():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        command = bridge.issue("writer", "smoke_docx", {})
        bridge.state.complete(ProbeResult(command.id, True, {"saved": True}, None))
        with pytest.raises(ProtocolError, match="Conflicting duplicate"):
            bridge.state.complete(ProbeResult(command.id, False, {}, {"code": "failed"}))
```

- [ ] **Step 2: Run the bridge tests and verify the module is missing**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_bridge.py -q`

Expected: FAIL during collection because `bridge.py` does not exist.

- [ ] **Step 3: Implement bridge state with per-component queues**

Implement `BridgeState` in `bridge.py` with:

```python
class BridgeState:
    def __init__(self):
        self._queues = {component: Queue() for component in COMPONENTS}
        self._registered: set[str] = set()
        self._results: dict[str, ProbeResult] = {}
        self._canceled: set[str] = set()
        self._last_seen: dict[str, float] = {}
        self._condition = Condition()

    def register(self, component: str) -> None:
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")
        with self._condition:
            self._registered.add(component)
            self._last_seen[component] = monotonic()
            self._condition.notify_all()

    def issue(self, component: str, method: str, params: Mapping[str, Any]) -> ProbeCommand:
        command = ProbeCommand.create(component, method, params)
        self._queues[component].put(command)
        return command

    def next(self, component: str, timeout: float = 10.0) -> Optional[ProbeCommand]:
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")
        with self._condition:
            self._last_seen[component] = monotonic()
        try:
            return self._queues[component].get(timeout=timeout)
        except Empty:
            return None

    def complete(self, result: ProbeResult) -> None:
        with self._condition:
            if result.id in self._canceled:
                return
            existing = self._results.get(result.id)
            if existing is not None and existing != result:
                raise ProtocolError(f"Conflicting duplicate result: {result.id}")
            self._results[result.id] = result
            self._condition.notify_all()

    def cancel(self, command_id: str) -> None:
        with self._condition:
            self._canceled.add(command_id)
            self._results.pop(command_id, None)
            self._condition.notify_all()

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        deadline = monotonic() + timeout
        with self._condition:
            while command_id not in self._results:
                remaining = deadline - monotonic()
                if remaining <= 0:
                    raise TimeoutError(f"Timed out waiting for result {command_id}")
                self._condition.wait(remaining)
            return self._results.pop(command_id)

    def wait_registered(self, expected: set[str], timeout: float) -> None:
        deadline = monotonic() + timeout
        with self._condition:
            while not expected <= self._registered:
                remaining = deadline - monotonic()
                if remaining <= 0:
                    missing = sorted(expected - self._registered)
                    raise TimeoutError(f"Timed out waiting for components: {missing}")
                self._condition.wait(remaining)

    def last_seen(self, component: str) -> Optional[float]:
        with self._condition:
            return self._last_seen.get(component)
```

Treat every authenticated long-poll request as the component heartbeat by updating `_last_seen`. Expose a read-only `last_seen(component)` method for diagnostics. Phase 0 cancellation is intentionally host-side: when a command times out, the runner calls `cancel(command_id)`, ignores a late result, and records `cancellation: mapped` because a synchronous WPS JSAPI call cannot be interrupted safely. Command IDs make completion idempotent: an identical duplicate result is accepted, while a conflicting duplicate raises `ProtocolError`.

Use `ThreadingHTTPServer(("127.0.0.1", 0), Handler)` and a handler factory closed over `state`, `token`, and `allowed_origins`. Implement only these routes:

```text
OPTIONS *                         -> 204 after origin validation
POST /v1/register                -> 204, body {"component": str}
GET  /v1/next?component=<name>   -> 200 command JSON or 204 after 10 seconds
POST /v1/result                  -> 204, ProbeResult JSON
GET  /health                     -> 200, {"status": "ok"}
```

For every route, validate `Origin` against the exact allow-list. Validate `Authorization` with `secrets.compare_digest` on GET and POST; the browser's `OPTIONS` preflight does not carry the bearer token and is accepted only for an allow-listed origin. Return JSON errors as `{"error":{"code":...,"message":...}}`; never echo the token. Add these response headers on accepted origins:

```python
self.send_header("Access-Control-Allow-Origin", origin)
self.send_header("Vary", "Origin")
self.send_header("Access-Control-Allow-Headers", "Authorization, Content-Type")
self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
```

Expose the state through `LoopbackBridge.state` for deterministic tests and add these delegates so callers never reach into handler internals:

```python
def issue(self, component: str, method: str, params: Mapping[str, Any]) -> ProbeCommand:
    return self.state.issue(component, method, params)

def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
    return self.state.wait_result(command_id, timeout)

def wait_registered(self, expected: set[str], timeout: float) -> None:
    self.state.wait_registered(expected, timeout)
```

`LoopbackBridge.__enter__()` starts `serve_forever()` in a daemon thread; `__exit__()` calls `shutdown()`, `server_close()`, and joins the thread. Generate the token with `secrets.token_urlsafe(32)`.

- [ ] **Step 4: Run focused bridge tests**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_bridge.py -q`

Expected: `6 passed`.

- [ ] **Step 5: Run the complete pure-Python suite**

Run: `.venv/bin/python -m pytest -q`

Expected: all existing tests plus the eleven new probe tests pass.

- [ ] **Step 6: Commit the bridge**

```bash
git add skills/WPSComposer/scripts/macos_probe/bridge.py tests/macos_probe/test_bridge.py
git commit -m "Add authenticated macOS probe bridge"
```

### Task 3: Pin the official WPS add-in toolchain and build reversible runtime profiles

**Files:**
- Create: `macos/wps-jsapi-probe/package.json`
- Create: `macos/wps-jsapi-probe/package-lock.json`
- Create: `skills/WPSComposer/scripts/macos_probe/runtime.py`
- Create: `tests/macos_probe/test_runtime.py`
- Modify: `.gitignore`

**Interfaces:**
- Consumes: `PathPolicy` and the add-in assets created in Task 4.
- Produces: `ProbeRuntime.prepare_profiles()`, `ProbeRuntime.start_servers()`, `ProbeRuntime.activate_components()`, `ProbeRuntime.restore_registration()`, and `read_wps_version()`.

- [ ] **Step 1: Add the pinned Node package manifest**

```json
{
  "name": "wpscomposer-macos-jsapi-probe",
  "version": "0.1.0",
  "private": true,
  "description": "Phase 0 Mac WPS JSAPI feasibility probe",
  "devDependencies": {
    "wps-jsapi-declare": "2.2.0",
    "wpsjs": "2.2.3"
  }
}
```

- [ ] **Step 2: Generate and inspect the lock file with the known-good Node runtime**

Run:

```bash
cd macos/wps-jsapi-probe
PATH=/Users/neomei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin:$PATH npm install --package-lock-only
npm ls --depth=0
```

Expected: `wpsjs@2.2.3` and `wps-jsapi-declare@2.2.0`, with no dependency resolution error. Do not commit `node_modules`.

- [ ] **Step 3: Write failing runtime tests**

```python
import json
from pathlib import Path

from skills.WPSComposer.scripts.macos_probe.runtime import (
    COMPONENT_CONFIG,
    RegistrationSnapshot,
    build_profile,
    find_node,
)


def test_component_config_uses_distinct_ports_and_wps_types():
    assert COMPONENT_CONFIG == {
        "writer": {"addon_type": "wps", "port": 3891, "script": "writer.js"},
        "presentation": {"addon_type": "wpp", "port": 3892, "script": "presentation.js"},
        "spreadsheet": {"addon_type": "et", "port": 3893, "script": "spreadsheet.js"},
    }


def test_build_profile_writes_runtime_config(tmp_path: Path):
    assets = tmp_path / "assets"
    assets.mkdir()
    for name in ("index.html", "manifest.xml", "ribbon.xml", "bridge-client.js", "writer.js"):
        (assets / name).write_text(name, encoding="utf-8")

    profile = build_profile(
        assets,
        tmp_path / "profiles",
        "writer",
        "http://127.0.0.1:45678",
        "secret-token",
    )

    package = json.loads((profile / "package.json").read_text())
    session = json.loads((profile / "session.json").read_text())
    assert package["addonType"] == "wps"
    assert session == {
        "bridgeUrl": "http://127.0.0.1:45678",
        "component": "writer",
        "token": "secret-token",
    }
    assert (profile / "component.js").read_text() == "writer.js"


def test_registration_snapshot_restores_existing_bytes(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>\n"
    publish.write_bytes(original)
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    publish.write_bytes(b"changed")
    snapshot.restore()
    assert publish.read_bytes() == original
    assert not (tmp_path / "recovery").exists()


def test_registration_snapshot_removes_probe_created_file(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    publish.write_bytes(b"probe")
    snapshot.restore()
    assert not publish.exists()
    assert not (tmp_path / "recovery").exists()


def test_find_node_honors_explicit_override(tmp_path: Path):
    node = tmp_path / "node"
    node.write_text("#!/bin/sh\nprintf 'v24.0.0\\n'\n")
    node.chmod(0o755)
    assert find_node(str(node)) == node.resolve()
```

- [ ] **Step 4: Run runtime tests and verify they fail**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_runtime.py -q`

Expected: FAIL during collection because `runtime.py` does not exist.

- [ ] **Step 5: Implement runtime profile creation and registration rollback**

Use these exact constants in `runtime.py`:

```python
WPS_APP = Path("/Applications/wpsoffice.app")
PUBLISH_XML = Path.home() / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
COMPONENT_CONFIG = {
    "writer": {"addon_type": "wps", "port": 3891, "script": "writer.js"},
    "presentation": {"addon_type": "wpp", "port": 3892, "script": "presentation.js"},
    "spreadsheet": {"addon_type": "et", "port": 3893, "script": "spreadsheet.js"},
}
```

`build_profile()` must create `<runtime>/profiles/<component>`, copy `index.html`, `manifest.xml`, `ribbon.xml`, and `bridge-client.js`, copy the selected component script as `component.js`, and write:

```python
package = {
    "name": f"wpscomposer-phase0-{component}",
    "version": "0.1.0",
    "private": True,
    "addonType": config["addon_type"],
}
session = {
    "bridgeUrl": bridge_url,
    "component": component,
    "token": token,
}
```

`RegistrationSnapshot.capture(path, recovery_dir)` stores `path.read_bytes()` when present and `None` otherwise. Before WPS registration changes, it creates `recovery_dir` with mode `0700`, writes `registration.json` containing the target path and original-existence flag, and writes the original bytes to `publish.xml.original` with mode `0600` when present. Print the absolute recovery directory before starting `wpsjs` so a hard-killed process can be repaired manually.

`restore()` uses `path.parent.mkdir(parents=True, exist_ok=True)`. For an existing original, write bytes to a same-directory temporary file, flush and `os.fsync()`, then replace the target atomically with `os.replace()`. When the file did not exist before the run, call `path.unlink(missing_ok=True)`. After verifying the restored bytes or absence, remove only the two recovery files and their now-empty recovery directory. The recovery copy therefore survives `SIGKILL` but is removed after normal cleanup.

`find_node(override)` checks the explicit argument, `WPSCOMPOSER_NODE`, then `shutil.which("node")`. For each candidate run `[candidate, "--version"]`, parse the leading major version, and require `major >= 20`; otherwise raise `RuntimeError("Node.js 20 or newer is required")`.

`find_wpsjs_cli(probe_root)` returns `<probe_root>/node_modules/wpsjs/src/index.js` and raises an actionable error instructing the user to run `npm install` when absent.

Append to `.gitignore`:

```gitignore
macos/wps-jsapi-probe/node_modules/
macos/wps-jsapi-probe/.runtime/
build/macos-phase0/
```

- [ ] **Step 6: Implement process and activation lifecycle**

`ProbeRuntime` must:

1. Reject non-macOS platforms and missing `/Applications/wpsoffice.app`.
2. Snapshot `PUBLISH_XML` to a printed, private recovery directory before starting any server.
3. Verify ports 3891, 3892, and 3893 are free by binding temporary sockets to `127.0.0.1`.
4. Spawn one process per profile, in Writer → Presentation → Spreadsheet order, and wait for each HTTP server before starting the next. Sequential startup prevents two `wpsjs` processes from reading and overwriting the same `publish.xml` concurrently. Use:

```python
[
    str(node),
    str(wpsjs_cli),
    "debug",
    "--server",
    "--port",
    str(config["port"]),
]
```

5. Set each process `cwd` to its profile and capture merged stdout/stderr in a per-component log file.
6. Poll `http://127.0.0.1:<port>/index.html` for at most 15 seconds.
7. Copy `wpsDemo.docx`, `wppDemo.pptx`, and `etDemo.xlsx` from `node_modules/wpsjs/src/lib/res` into the run directory.
8. Activate components with `open -a /Applications/wpsoffice.app <fixture>`; never invoke Accessibility or keyboard automation.
9. On cleanup, terminate only the three child Node processes, wait five seconds, kill a child only if it did not terminate, and restore `publish.xml` in `finally`.

`wpsjs 2.2.3` selects its `debug_publish` path on Darwin; that implementation still launches WPS even when `--server` is supplied. Treat those launches as expected and harmless, then open the three fixture paths explicitly so every component add-in loads. Do not claim that `--server` suppresses WPS on this version.

Read WPS version using Python's plist parser so the path is handled without a shell domain lookup:

```python
with (WPS_APP / "Contents/Info.plist").open("rb") as stream:
    value = plistlib.load(stream)["CFBundleShortVersionString"]
return str(value)
```

- [ ] **Step 7: Run runtime and full tests**

Run:

```bash
.venv/bin/python -m pytest tests/macos_probe/test_runtime.py -q
.venv/bin/python -m pytest -q
```

Expected: runtime tests report `5 passed`; full suite remains green.

- [ ] **Step 8: Commit the reproducible runtime**

```bash
git add .gitignore macos/wps-jsapi-probe skills/WPSComposer/scripts/macos_probe/runtime.py tests/macos_probe/test_runtime.py
git commit -m "Add reversible WPS JSAPI probe runtime"
```

### Task 4: Add the shared WPS add-in shell and authenticated bridge client

**Files:**
- Create: `macos/wps-jsapi-probe/addin/index.html`
- Create: `macos/wps-jsapi-probe/addin/manifest.xml`
- Create: `macos/wps-jsapi-probe/addin/ribbon.xml`
- Create: `macos/wps-jsapi-probe/addin/bridge-client.js`
- Create: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: `/v1/register`, `/v1/next`, and `/v1/result` from Task 2 and `session.json` from Task 3.
- Produces: global `OnAddinLoad()` and `window.WPSComposerProbe.handleCommand` contract.

- [ ] **Step 1: Write failing static asset tests**

```python
from pathlib import Path

ROOT = Path("macos/wps-jsapi-probe/addin")


def test_manifest_and_ribbon_are_minimal_and_load_probe():
    manifest = (ROOT / "manifest.xml").read_text()
    ribbon = (ROOT / "ribbon.xml").read_text()
    assert "<Name>WPSComposerPhase0</Name>" in manifest
    assert 'onLoad="OnAddinLoad"' in ribbon
    assert "<button" not in ribbon


def test_html_loads_shared_client_before_component():
    html = (ROOT / "index.html").read_text()
    assert html.index("bridge-client.js") < html.index("component.js")


def test_bridge_client_has_no_dynamic_code_execution():
    source = (ROOT / "bridge-client.js").read_text()
    assert "Authorization" in source
    assert "/v1/register" in source
    assert "/v1/next" in source
    assert "/v1/result" in source
    assert "eval(" not in source
    assert "new Function" not in source
```

- [ ] **Step 2: Run the asset tests and verify missing files**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py -q`

Expected: three failures caused by missing add-in assets.

- [ ] **Step 3: Add the manifest, HTML, and invisible ribbon**

`manifest.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<JsPlugin>
  <ApiVersion>1.0.0</ApiVersion>
  <Name>WPSComposerPhase0</Name>
  <Description>WPSComposer macOS JSAPI feasibility probe</Description>
</JsPlugin>
```

`index.html`:

```html
<!doctype html>
<html>
  <head><meta charset="utf-8"></head>
  <body>
    <script src="./bridge-client.js"></script>
    <script src="./component.js"></script>
  </body>
</html>
```

`ribbon.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="OnAddinLoad">
  <ribbon startFromScratch="false"><tabs/></ribbon>
</customUI>
```

- [ ] **Step 4: Implement the exact client loop**

`bridge-client.js` must expose no arbitrary dispatch. Use:

```javascript
(function () {
  "use strict";

  let started = false;

  async function readJson(response) {
    const text = await response.text();
    return text ? JSON.parse(text) : null;
  }

  async function request(session, path, options) {
    const headers = Object.assign({}, options.headers || {}, {
      Authorization: `Bearer ${session.token}`,
      "Content-Type": "application/json"
    });
    const response = await fetch(`${session.bridgeUrl}${path}`, Object.assign({}, options, {headers}));
    if (!response.ok && response.status !== 204) {
      throw new Error(`Bridge ${path} failed with HTTP ${response.status}`);
    }
    return {status: response.status, body: await readJson(response)};
  }

  async function sendResult(session, result) {
    await request(session, "/v1/result", {method: "POST", body: JSON.stringify(result)});
  }

  async function run() {
    const sessionResponse = await fetch("./session.json", {cache: "no-store"});
    if (!sessionResponse.ok) throw new Error("session.json is unavailable");
    const session = await sessionResponse.json();
    await request(session, "/v1/register", {
      method: "POST",
      body: JSON.stringify({component: session.component})
    });

    while (true) {
      const next = await request(
        session,
        `/v1/next?component=${encodeURIComponent(session.component)}`,
        {method: "GET"}
      );
      if (next.status === 204) continue;
      const command = next.body;
      try {
        const value = await window.WPSComposerProbe.handleCommand(command);
        await sendResult(session, {id: command.id, ok: true, value, error: null});
      } catch (error) {
        await sendResult(session, {
          id: command.id,
          ok: false,
          value: {},
          error: {
            code: "JSAPI_COMMAND_FAILED",
            message: String(error && error.message ? error.message : error),
            stack: String(error && error.stack ? error.stack : "")
          }
        });
      }
    }
  }

  window.OnAddinLoad = function () {
    if (!started) {
      started = true;
      run().catch(function (error) {
        console.error("WPSComposer Phase 0 add-in failed", error);
      });
    }
    return true;
  };
}());
```

- [ ] **Step 5: Run asset tests**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py -q`

Expected: `3 passed`.

- [ ] **Step 6: Commit the shared add-in shell**

```bash
git add macos/wps-jsapi-probe/addin tests/macos_probe/test_addin_assets.py
git commit -m "Add WPS probe add-in shell"
```

### Task 5: Implement Writer DOCX and standalone PDF smoke commands

**Files:**
- Create: `macos/wps-jsapi-probe/addin/writer.js`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: `window.Application`, command methods `probe_capabilities`, `smoke_docx`, and `smoke_pdf`.
- Produces: DOCX saved with WPS format `12`, PDF exported with WPS format `17`, and per-operation capability records.

- [ ] **Step 1: Extend the static test with Writer invariants**

```python
def test_writer_uses_wps_save_and_pdf_export():
    source = (ROOT / "writer.js").read_text()
    assert "Application.Documents.Add" in source
    assert "Tables.Add" in source
    assert "InlineShapes.AddPicture" in source
    assert "SaveAs2(outputPath, 12)" in source
    assert "ExportAsFixedFormat(outputPath, 17" in source
    assert '"smoke_docx"' in source
    assert '"smoke_pdf"' in source
```

- [ ] **Step 2: Run the focused test and verify `writer.js` is missing**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py::test_writer_uses_wps_save_and_pdf_export -q`

Expected: FAIL with `FileNotFoundError`.

- [ ] **Step 3: Implement Writer capability classification and document creation**

Create a closed method table; never index `Application` using command input:

```javascript
(function () {
  "use strict";
  const capabilities = {};

  function record(name, classification, detail) {
    capabilities[name] = {classification, detail: detail || ""};
  }

  function attempt(name, action, classification) {
    try {
      const value = action();
      record(name, classification || "native", "");
      return value;
    } catch (error) {
      record(name, "unsupported", String(error && error.message ? error.message : error));
      return null;
    }
  }

  function typeParagraph(selection, text, size, bold) {
    selection.Font.Name = "PingFang SC";
    selection.Font.NameAscii = "Helvetica Neue";
    selection.Font.Size = size;
    selection.Font.Bold = bold;
    selection.TypeText(text);
    selection.TypeParagraph();
  }

  function addTable(document, selection) {
    const table = document.Tables.Add(selection.Range, 3, 2);
    const values = [["项目", "状态"], ["Writer JSAPI", "通过"], ["PDF 导出", "通过"]];
    for (let row = 1; row <= values.length; row += 1) {
      for (let column = 1; column <= values[row - 1].length; column += 1) {
        table.Cell(row, column).Range.Text = values[row - 1][column - 1];
      }
    }
    table.Rows.Item(1).Range.Font.Bold = true;
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
  }

  function addImage(document, selection, imagePath) {
    selection.SetRange(document.Content.End, document.Content.End);
    const image = document.InlineShapes.AddPicture(imagePath, false, true, selection.Range);
    image.Width = 72;
    image.Height = 72;
    selection.SetRange(document.Content.End, document.Content.End);
    selection.TypeParagraph();
  }

  function buildDocument(mode, imagePath) {
    const document = Application.Documents.Add();
    const selection = Application.Selection;
    typeParagraph(selection, mode === "pdf" ? "WPSComposer Obsidian Reading" : "WPSComposer Writer Smoke", 24, true);
    typeParagraph(selection, "macOS WPS JSAPI 可行性验证", 12, false);
    typeParagraph(selection, "## 能力概览", 18, true);
    typeParagraph(selection, "正文由 WPS Writer 在 macOS 上排版并序列化。", 11, false);
    if (mode === "pdf") {
      typeParagraph(selection, "```python", 10, true);
      typeParagraph(selection, "print('Rendered by Mac WPS')", 10, false);
      typeParagraph(selection, "```", 10, true);
      typeParagraph(selection, "> [!TIP] 这是 Obsidian 风格提示块。", 11, false);
      typeParagraph(selection, "☑ WPS 排版", 11, false);
      typeParagraph(selection, "☐ Phase 1 集成", 11, false);
    }
    addTable(document, selection);
    addImage(document, selection, imagePath);
    return document;
  }

  function saveDocx(params) {
    const outputPath = params.outputPath;
    const document = buildDocument("docx", params.imagePath);
    try {
      document.SaveAs2(outputPath, 12);
      record("writer.save_docx", "native", "Document.SaveAs2 format 12");
      return {path: outputPath, capabilities};
    } finally {
      document.Close(0);
    }
  }

  function savePdf(params) {
    const outputPath = params.outputPath;
    const document = buildDocument("pdf", params.imagePath);
    try {
      document.ExportAsFixedFormat(outputPath, 17, false, 0, 0);
      record("writer.export_pdf", "native", "Document.ExportAsFixedFormat format 17");
      return {path: outputPath, preset: "obsidian-reading", capabilities};
    } finally {
      document.Close(0);
    }
  }

  function probe() {
    attempt("writer.documents_add", function () { return Application.Documents.Count; });
    attempt("writer.font_enumeration", function () { return Application.FontNames.Count; });
    attempt("writer.template_enumeration", function () { return Application.Templates.Count; });
    record("writer.tables", "native", "Document.Tables.Add");
    record("writer.images", "native", "Document.InlineShapes.AddPicture");
    return {component: "writer", applicationName: Application.Name, capabilities};
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_docx": saveDocx,
    "smoke_pdf": savePdf
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) throw new Error(`Writer rejects method ${command.method}`);
      return handler(command.params || {});
    }
  };
}());
```

- [ ] **Step 4: Run the Writer static test and full asset tests**

Run:

```bash
.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py::test_writer_uses_wps_save_and_pdf_export -q
.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py -q
```

Expected: focused test passes and the asset suite reports `4 passed`.

- [ ] **Step 5: Commit the Writer probe**

```bash
git add macos/wps-jsapi-probe/addin/writer.js tests/macos_probe/test_addin_assets.py
git commit -m "Add Mac WPS Writer smoke commands"
```

### Task 6: Implement Presentation PPTX smoke command

**Files:**
- Create: `macos/wps-jsapi-probe/addin/presentation.js`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: `window.Application`, `probe_capabilities`, and `smoke_pptx`.
- Produces: a WPS-created PPTX with title, content, image, and table slides, saved with format `24`.

- [ ] **Step 1: Add a failing Presentation contract test**

```python
def test_presentation_uses_wps_shapes_and_save():
    source = (ROOT / "presentation.js").read_text()
    assert "Application.Presentations.Add" in source
    assert "Shapes.AddTextbox" in source
    assert "Shapes.AddPicture" in source
    assert "Shapes.AddTable" in source
    assert "SaveAs(outputPath, 24)" in source
    assert '"smoke_pptx"' in source
```

- [ ] **Step 2: Run the test and verify the asset is absent**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py::test_presentation_uses_wps_shapes_and_save -q`

Expected: FAIL with `FileNotFoundError`.

- [ ] **Step 3: Implement a fixed, typed Presentation renderer**

Use `ppLayoutBlank = 12`, `msoTextOrientationHorizontal = 1`, `msoFalse = 0`, `msoTrue = -1`, and this closed command handler:

```javascript
(function () {
  "use strict";
  const capabilities = {};

  function text(slide, value, left, top, width, height, size, bold) {
    const shape = slide.Shapes.AddTextbox(1, left, top, width, height);
    shape.TextFrame.TextRange.Text = value;
    shape.TextFrame.TextRange.Font.Name = "PingFang SC";
    shape.TextFrame.TextRange.Font.Size = size;
    shape.TextFrame.TextRange.Font.Bold = bold ? -1 : 0;
    return shape;
  }

  function tableSlide(presentation) {
    const slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12);
    text(slide, "能力矩阵", 36, 24, 600, 40, 26, true);
    const shape = slide.Shapes.AddTable(3, 2, 72, 100, 576, 240);
    const values = [["能力", "结果"], ["文本与图片", "native"], ["表格与保存", "native"]];
    for (let row = 1; row <= 3; row += 1) {
      for (let column = 1; column <= 2; column += 1) {
        shape.Table.Cell(row, column).Shape.TextFrame.TextRange.Text = values[row - 1][column - 1];
      }
    }
  }

  function savePptx(params) {
    const outputPath = params.outputPath;
    const presentation = Application.Presentations.Add(-1);
    try {
      let slide = presentation.Slides.Add(1, 12);
      text(slide, "WPSComposer macOS", 48, 80, 620, 80, 32, true);
      text(slide, "WPS Presentation JSAPI Phase 0", 48, 180, 620, 48, 18, false);

      slide = presentation.Slides.Add(2, 12);
      text(slide, "内容与图片", 36, 24, 600, 40, 26, true);
      text(slide, "• WPS 创建演示文稿\n• JSAPI 添加形状\n• WPS 保存 PPTX", 48, 100, 300, 240, 20, false);
      slide.Shapes.AddPicture(params.imagePath, 0, -1, 400, 120, 160, 160);
      tableSlide(presentation);

      presentation.SaveAs(outputPath, 24);
      capabilities["presentation.create"] = {classification: "native", detail: "Presentations.Add"};
      capabilities["presentation.text"] = {classification: "native", detail: "Shapes.AddTextbox"};
      capabilities["presentation.image"] = {classification: "native", detail: "Shapes.AddPicture"};
      capabilities["presentation.table"] = {classification: "native", detail: "Shapes.AddTable"};
      capabilities["presentation.save_pptx"] = {classification: "native", detail: "Presentation.SaveAs format 24"};
      return {path: outputPath, capabilities};
    } finally {
      presentation.Close();
    }
  }

  function probe() {
    return {
      component: "presentation",
      applicationName: Application.Name,
      capabilities: {
        "presentation.presentations": {classification: "native", detail: String(Application.Presentations.Count)},
        "presentation.slides": {classification: "native", detail: "Slides.Add"},
        "presentation.shapes": {classification: "native", detail: "Shapes collection"}
      }
    };
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_pptx": savePptx
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) throw new Error(`Presentation rejects method ${command.method}`);
      return handler(command.params || {});
    }
  };
}());
```

- [ ] **Step 4: Run the Presentation and full asset suites**

Run:

```bash
.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py::test_presentation_uses_wps_shapes_and_save -q
.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py -q
```

Expected: focused test passes and the asset suite reports `5 passed`.

- [ ] **Step 5: Commit the Presentation probe**

```bash
git add macos/wps-jsapi-probe/addin/presentation.js tests/macos_probe/test_addin_assets.py
git commit -m "Add Mac WPS Presentation smoke command"
```

### Task 7: Implement Spreadsheet XLSX smoke command

**Files:**
- Create: `macos/wps-jsapi-probe/addin/spreadsheet.js`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: `window.Application`, `probe_capabilities`, and `smoke_xlsx`.
- Produces: a WPS-created XLSX with values, formulas, formatting, and a chart, saved with format `51`.

- [ ] **Step 1: Add a failing Spreadsheet contract test**

```python
def test_spreadsheet_uses_wps_cells_chart_and_save():
    source = (ROOT / "spreadsheet.js").read_text()
    assert "Application.Workbooks.Add" in source
    assert '.Formula = "=B2*C2"' in source
    assert "ChartObjects().Add" in source
    assert "SetSourceData" in source
    assert "SaveAs(outputPath, 51)" in source
    assert '"smoke_xlsx"' in source
```

- [ ] **Step 2: Run the focused test and verify the file is absent**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py::test_spreadsheet_uses_wps_cells_chart_and_save -q`

Expected: FAIL with `FileNotFoundError`.

- [ ] **Step 3: Implement the Spreadsheet command**

```javascript
(function () {
  "use strict";

  function saveXlsx(params) {
    const outputPath = params.outputPath;
    const workbook = Application.Workbooks.Add();
    try {
      const sheet = workbook.Worksheets.Item(1);
      sheet.Name = "Phase0";
      const values = [
        ["项目", "数量", "单价", "金额"],
        ["Writer", 2, 120, null],
        ["Presentation", 3, 160, null],
        ["Spreadsheet", 4, 200, null]
      ];
      for (let row = 1; row <= values.length; row += 1) {
        for (let column = 1; column <= values[row - 1].length; column += 1) {
          if (values[row - 1][column - 1] !== null) {
            sheet.Cells.Item(row, column).Value2 = values[row - 1][column - 1];
          }
        }
      }
      sheet.Range("D2").Formula = "=B2*C2";
      sheet.Range("D2:D4").FillDown();
      sheet.Range("A1:D1").Font.Bold = true;
      sheet.Range("A1:D1").Interior.Color = 12611584;
      sheet.Range("A1:D4").Borders.LineStyle = 1;
      sheet.UsedRange.Columns.AutoFit();

      const chartObject = sheet.ChartObjects().Add(320, 40, 480, 260);
      chartObject.Chart.SetSourceData(sheet.Range("A1:D4"));
      chartObject.Chart.ChartType = 51;
      chartObject.Chart.HasTitle = true;
      chartObject.Chart.ChartTitle.Text = "WPSComposer Phase 0";

      workbook.SaveAs(outputPath, 51);
      return {
        path: outputPath,
        capabilities: {
          "spreadsheet.create": {classification: "native", detail: "Workbooks.Add"},
          "spreadsheet.values": {classification: "native", detail: "Cells.Item.Value2"},
          "spreadsheet.formulas": {classification: "native", detail: "Range.Formula and FillDown"},
          "spreadsheet.formatting": {classification: "native", detail: "Font, Interior, Borders, AutoFit"},
          "spreadsheet.chart": {classification: "native", detail: "ChartObjects.Add and SetSourceData"},
          "spreadsheet.save_xlsx": {classification: "native", detail: "Workbook.SaveAs format 51"}
        }
      };
    } finally {
      workbook.Close(false);
    }
  }

  function probe() {
    return {
      component: "spreadsheet",
      applicationName: Application.Name,
      capabilities: {
        "spreadsheet.workbooks": {classification: "native", detail: String(Application.Workbooks.Count)},
        "spreadsheet.range": {classification: "native", detail: "Worksheet.Range"},
        "spreadsheet.charts": {classification: "native", detail: "Worksheet.ChartObjects"}
      }
    };
  }

  const handlers = {
    "probe_capabilities": function () { return probe(); },
    "smoke_xlsx": saveXlsx
  };

  window.WPSComposerProbe = {
    handleCommand: function (command) {
      const handler = handlers[command.method];
      if (!handler) throw new Error(`Spreadsheet rejects method ${command.method}`);
      return handler(command.params || {});
    }
  };
}());
```

- [ ] **Step 4: Run static and full Python tests**

Run:

```bash
.venv/bin/python -m pytest tests/macos_probe/test_addin_assets.py -q
.venv/bin/python -m pytest -q
```

Expected: asset suite reports `6 passed`; complete suite remains green.

- [ ] **Step 5: Commit the Spreadsheet probe**

```bash
git add macos/wps-jsapi-probe/addin/spreadsheet.js tests/macos_probe/test_addin_assets.py
git commit -m "Add Mac WPS Spreadsheet smoke command"
```

### Task 8: Orchestrate all components and write a deterministic Phase 0 report

**Files:**
- Create: `skills/WPSComposer/scripts/macos_probe/runner.py`
- Create: `skills/WPSComposer/scripts/macos_probe/__main__.py`
- Create: `tests/macos_probe/test_runner.py`

**Interfaces:**
- Consumes: `LoopbackBridge`, `ProbeRuntime`, `PathPolicy`, and the five typed command methods.
- Produces: `run_phase0(output_dir, node=None, timeout=90) -> Path`, four artifacts, `capabilities.json`, and `phase0-report.json`.

- [ ] **Step 1: Write failing artifact and orchestration tests**

```python
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.macos_probe.bridge import LoopbackBridge
from skills.WPSComposer.scripts.macos_probe.models import PathPolicy, ProbeResult
from skills.WPSComposer.scripts.macos_probe.runner import (
    _execute_commands,
    ArtifactError,
    validate_artifact,
)


def test_validate_native_zip_artifacts(tmp_path: Path):
    for suffix in ("docx", "pptx", "xlsx"):
        path = tmp_path / f"smoke.{suffix}"
        path.write_bytes(b"PK\x03\x04" + b"x" * 2048)
        validate_artifact(path, suffix)


def test_validate_pdf_signature(tmp_path: Path):
    path = tmp_path / "smoke.pdf"
    path.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
    validate_artifact(path, "pdf")


def test_validate_artifact_rejects_empty_or_wrong_signature(tmp_path: Path):
    path = tmp_path / "bad.docx"
    path.write_bytes(b"not a package")
    with pytest.raises(ArtifactError, match="invalid DOCX signature"):
        validate_artifact(path, "docx")


def fake_component(bridge, component, command_count):
    bridge.state.register(component)
    for _ in range(command_count):
        command = bridge.state.next(component, 2)
        assert command is not None
        if command.method == "probe_capabilities":
            value = {
                "component": component,
                "applicationName": f"Fake {component}",
                "capabilities": {
                    f"{component}.probe": {"classification": "native", "detail": "fake"}
                },
            }
        else:
            path = Path(command.params["outputPath"])
            signature = b"%PDF-1.7\n" if path.suffix == ".pdf" else b"PK\x03\x04"
            path.write_bytes(signature + b"x" * 2048)
            value = {"path": str(path), "capabilities": {}}
            if path.suffix == ".pdf":
                value["preset"] = "obsidian-reading"
        bridge.state.complete(ProbeResult(command.id, True, value, None))


def test_execute_commands_keeps_four_outputs_independent(tmp_path: Path):
    output = tmp_path / "output"
    output.mkdir()
    image = tmp_path / "fixture.png"
    image.write_bytes(b"png")
    policy = PathPolicy((tmp_path, output))
    origins = {
        "http://127.0.0.1:3891",
        "http://127.0.0.1:3892",
        "http://127.0.0.1:3893",
    }

    with LoopbackBridge(origins) as bridge, ThreadPoolExecutor(max_workers=3) as pool:
        futures = (
            pool.submit(fake_component, bridge, "writer", 3),
            pool.submit(fake_component, bridge, "presentation", 2),
            pool.submit(fake_component, bridge, "spreadsheet", 2),
        )
        bridge.wait_registered({"writer", "presentation", "spreadsheet"}, 2)
        report = _execute_commands(bridge, policy, output, image, timeout=2)
        for future in futures:
            future.result(timeout=2)

    assert [item["format"] for item in report["artifacts"]] == [
        "docx",
        "pptx",
        "xlsx",
        "pdf",
    ]
    assert report["status"] == "passed"
    assert report["pdfPreset"] == "obsidian-reading"
    assert "temporaryDocx" not in report
```

- [ ] **Step 2: Run runner tests and verify the module is missing**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_runner.py -q`

Expected: FAIL during collection because `runner.py` does not exist.

- [ ] **Step 3: Implement strict artifact validation**

```python
class ArtifactError(RuntimeError):
    pass


def validate_artifact(path: Path, format_name: str) -> None:
    if not path.is_file() or path.stat().st_size < 1024:
        raise ArtifactError(f"{format_name.upper()} output is missing or too small: {path}")
    signature = path.read_bytes()[:8]
    expected = b"%PDF-" if format_name == "pdf" else b"PK\x03\x04"
    if not signature.startswith(expected):
        raise ArtifactError(f"invalid {format_name.upper()} signature: {path}")
```

Generate a small local PNG in the run directory from this fixed base64 payload:

```python
PNG_FIXTURE = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8"
    "/x8AAusB9Y9Zl1sAAAAASUVORK5CYII="
)
```

- [ ] **Step 4: Implement command ordering and independent results**

Implement `_execute_commands(bridge, policy, output_dir, image_path, timeout) -> dict[str, Any]`. After all clients register, issue capability probes to each component and wait for each result. Then issue the outputs in this exact order and validate each before continuing:

```python
commands = (
    ("writer", "smoke_docx", "smoke.docx", "docx"),
    ("presentation", "smoke_pptx", "smoke.pptx", "pptx"),
    ("spreadsheet", "smoke_xlsx", "smoke.xlsx", "xlsx"),
    ("writer", "smoke_pdf", "smoke.pdf", "pdf"),
)
```

For every command, validate both `outputPath` and `imagePath` through `PathPolicy` before enqueueing. Wrap `wait_result()` so `TimeoutError` calls `bridge.state.cancel(command.id)` and records `cancellation` as `mapped`. A failed command must be recorded with its component, method, structured JSAPI error, and log path. Continue with other output components when safe so the report identifies partial capability; return a failing exit status unless all four artifacts validate.

Write `capabilities.json` with:

```json
{
  "schemaVersion": 1,
  "platform": "macOS",
  "wpsVersion": "<runtime value>",
  "components": {
    "writer": {},
    "presentation": {},
    "spreadsheet": {}
  }
}
```

Write `phase0-report.json` with:

```json
{
  "schemaVersion": 1,
  "status": "passed",
  "backend": "mac-wps-jsapi-probe",
  "platform": "macOS",
  "wpsVersion": "12.1.26035",
  "wpsjsVersion": "2.2.3",
  "pdfPreset": "obsidian-reading",
  "artifacts": [
    {"format": "docx", "path": "<absolute path>", "size": 0},
    {"format": "pptx", "path": "<absolute path>", "size": 0},
    {"format": "xlsx", "path": "<absolute path>", "size": 0},
    {"format": "pdf", "path": "<absolute path>", "size": 0}
  ],
  "failures": [],
  "timingsMs": {}
}
```

Replace the example version, sizes, paths, status, failures, and timings with runtime values before writing. Serialize JSON with `ensure_ascii=False`, `indent=2`, and a trailing newline.

- [ ] **Step 5: Add the command-line entrypoint**

`__main__.py` must accept:

```text
--output-dir PATH     default: build/macos-phase0
--node PATH           optional explicit Node executable
--timeout SECONDS     default: 90
```

Call `run_phase0()`, print only the absolute report path on success, print a concise error plus the preserved report path on failure, and exit `0` only when report status is `passed`.

- [ ] **Step 6: Run the runner and full suites**

Run:

```bash
.venv/bin/python -m pytest tests/macos_probe/test_runner.py -q
.venv/bin/python -m pytest -q
```

Expected: runner tests pass; the full suite remains green without launching WPS because real integration is not part of default pytest.

- [ ] **Step 7: Commit orchestration and reporting**

```bash
git add skills/WPSComposer/scripts/macos_probe/__main__.py skills/WPSComposer/scripts/macos_probe/runner.py tests/macos_probe/test_runner.py
git commit -m "Add Mac WPS Phase 0 probe runner"
```

### Task 9: Run the installed-Mac feasibility gate and document exact evidence

**Files:**
- Create: `docs/macos-phase0.md`
- Modify: `README.md`
- Modify: `docs/superpowers/specs/2026-07-16-macos-wps-jsapi-backend-design.md`
- Runtime output only: `build/macos-phase0/*` (gitignored)

**Interfaces:**
- Consumes: the complete Phase 0 runner and the installed Mac WPS.
- Produces: an evidence-backed pass/fail decision and a capability-gap table that gates Phase 1.

- [ ] **Step 1: Install the pinned probe dependencies**

Run:

```bash
cd macos/wps-jsapi-probe
PATH=/Users/neomei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin:$PATH npm ci
```

Expected: install completes with `wpsjs@2.2.3`; `node_modules` remains ignored.

- [ ] **Step 2: Run the real Phase 0 probe**

Run from repository root:

```bash
.venv/bin/python -m skills.WPSComposer.scripts.macos_probe \
  --node /Users/neomei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node \
  --output-dir build/macos-phase0 \
  --timeout 120
```

Expected on success: WPS opens the three official fixture documents, the add-ins create and close only their own smoke documents, the command exits `0`, and prints the absolute path to `phase0-report.json`. If WPS displays a first-run or account modal, dismiss it manually and rerun; do not add UI automation.

- [ ] **Step 3: Verify artifact structure independently**

Run:

```bash
python3 - <<'PY'
from pathlib import Path
from zipfile import ZipFile

root = Path("build/macos-phase0")
required = {
    "smoke.docx": "word/document.xml",
    "smoke.pptx": "ppt/presentation.xml",
    "smoke.xlsx": "xl/workbook.xml",
}
for name, member in required.items():
    with ZipFile(root / name) as package:
        assert member in package.namelist(), (name, member)
pdf = (root / "smoke.pdf").read_bytes()
assert pdf.startswith(b"%PDF-") and len(pdf) > 1024
print("Mac WPS Phase 0 artifact structure: PASS")
PY
```

Expected: `Mac WPS Phase 0 artifact structure: PASS`.

- [ ] **Step 4: Verify registration rollback**

Before and after one additional probe run, calculate:

```bash
registration="$HOME/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml"
if test -f "$registration"; then shasum -a 256 "$registration"; else printf 'ABSENT\n'; fi
```

Expected: the hashes match. If the file did not exist before the run, it must not exist afterward.

- [ ] **Step 5: Write the evidence document**

Document:

- tested macOS version, WPS version, Node version, `wpsjs` version, and timestamp;
- the exact run command and output report path;
- a four-row artifact table with format, absolute path, size, WPS API used, and validation result;
- every capability entry classified as `native`, `mapped`, or `unsupported`, copying error text from `capabilities.json` without inventing support;
- the registration rollback hash result;
- recovery steps: stop only probe Node processes, use the recovery directory printed before registration to atomically copy `publish.xml.original` back (or remove the probe-created file when `registration.json` says `existed: false`), and reopen WPS;
- an explicit Phase 1 decision: proceed only if all four outputs validate and Writer template/font resolution has at least one supported path.

Add to README under platform support:

```markdown
- macOS: WPS JSAPI backend is in Phase 0 feasibility validation. See [macOS Phase 0](docs/macos-phase0.md) for tested versions, capabilities, and current limitations.
```

Update the design document's Phase 0 section with the actual result date and link to `docs/macos-phase0.md`; do not change its approved architecture.

- [ ] **Step 6: Run final verification before claiming success**

Run:

```bash
.venv/bin/python -m pytest -q
git diff --check
git status --short
```

Expected: all tests pass, `git diff --check` is silent, and status contains only the intended documentation changes before commit. Confirm `build/macos-phase0`, `node_modules`, and WPS registration files are not staged.

- [ ] **Step 7: Commit Phase 0 evidence**

```bash
git add README.md docs/macos-phase0.md docs/superpowers/specs/2026-07-16-macos-wps-jsapi-backend-design.md
git commit -m "Document Mac WPS JSAPI feasibility result"
```

## Reference Evidence Used by This Plan

- `wpsjs 2.2.3` package source uses `/Applications/wpsoffice.app` on Darwin and registers online add-ins at `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml`.
- WPS Writer exposes `Documents.Add`, `Document.SaveAs2`, `Document.ExportAsFixedFormat`, `Tables.Add`, and `InlineShapes.AddPicture`.
- WPS Presentation exposes `Presentations.Add`, `Slides.Add`, `Shapes.AddTextbox`, `Shapes.AddPicture`, `Shapes.AddTable`, and `Presentation.SaveAs`.
- WPS Spreadsheet exposes `Workbooks.Add`, range values/formulas/formatting, `ChartObjects.Add`, and `Workbook.SaveAs`.
- Official API references:
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/generate-the-first-wps-addin
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/obj
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/member/SaveAs2
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/InlineShapes/member/AddPicture
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/Presentations/member/Add
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/Shapes/member/AddTable
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wpp/Presentation/member/SaveAs
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Application/member/Range
  - https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/et/Workbook/member/SaveAs
  - https://www.npmjs.com/package/wpsjs?activeTab=versions
