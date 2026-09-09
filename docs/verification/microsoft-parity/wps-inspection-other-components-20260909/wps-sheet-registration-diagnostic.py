from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys
import time
import traceback
from urllib.parse import urlsplit


ROOT = Path("/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description")
SOURCE = ROOT / "docs/verification/microsoft-parity/installed-audit-04/sheet/edited.xlsx"
OUTPUT = Path("/tmp/wps-sheet-registration-diagnostic.json")
sys.path.insert(0, str(ROOT))

from skills.WPSComposer import inspect  # noqa: E402
from skills.WPSComposer.scripts.macos_probe import bridge as bridge_module  # noqa: E402
from skills.WPSComposer.scripts.macos_probe import profile_server  # noqa: E402
from skills.WPSComposer.scripts.macos_probe import runtime as runtime_module  # noqa: E402


events: list[dict[str, object]] = []
started = time.monotonic()


def record(event: str, **details: object) -> None:
    events.append({"seconds": round(time.monotonic() - started, 4), "event": event, **details})


original_get = profile_server._ProfileRequestHandler.do_GET
original_head = profile_server._ProfileRequestHandler.do_HEAD
original_claim = bridge_module.BridgeState.claim_session
original_register = bridge_module.BridgeState.register
original_activate = runtime_module.ProbeRuntime.activate_component


def audited_get(self):
    record("profile_request", method="GET", path=urlsplit(self.path).path)
    return original_get(self)


def audited_head(self):
    record("profile_request", method="HEAD", path=urlsplit(self.path).path)
    return original_head(self)


def audited_claim(self, component, client_id, capability):
    record("bridge_session_claim", component=component)
    return original_claim(self, component, client_id, capability)


def audited_register(self, component, session_nonce):
    record("bridge_register", component=component)
    return original_register(self, component, session_nonce)


def audited_activate(self, component, **kwargs):
    record(
        "activation_start",
        component=component,
        existing_pids=sorted(runtime_module.list_wps_pids(self.wps_app)),
    )
    result = original_activate(self, component, **kwargs)
    record(
        "activation_return",
        component=component,
        fixture=str(result),
        current_pids=sorted(runtime_module.list_wps_pids(self.wps_app)),
    )
    return result


profile_server._ProfileRequestHandler.do_GET = audited_get
profile_server._ProfileRequestHandler.do_HEAD = audited_head
bridge_module.BridgeState.claim_session = audited_claim
bridge_module.BridgeState.register = audited_register
runtime_module.ProbeRuntime.activate_component = audited_activate

report = {
    "source": str(SOURCE),
    "source_sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
    "events": events,
    "ok": False,
}
try:
    inspect(SOURCE, engine="wps", timeout=50)
    report["ok"] = True
except BaseException as exc:
    report["error_type"] = type(exc).__name__
    report["error"] = str(exc)
    report["traceback"] = traceback.format_exc()
finally:
    report["events"] = events
    OUTPUT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    print(OUTPUT)
