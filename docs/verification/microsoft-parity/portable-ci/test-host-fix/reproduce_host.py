"""Run scoped tests with host dependencies controlled; no native apps are invoked."""
from __future__ import annotations

import builtins
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[5]
sys.path.insert(0, str(ROOT))
if __name__ == "__main__":
    HOST = sys.argv[1]
    from skills.WPSComposer.scripts import orchestrator
    from skills.WPSComposer.scripts.macos_probe import longform_evidence

    # Change only modules under test, not Python/pytest's own platform.
    orchestrator.sys = SimpleNamespace(platform=HOST)

    def missing_wps_version():
        raise FileNotFoundError("simulated host has no WPS app bundle")

    longform_evidence.read_wps_version = missing_wps_version
    real_import = builtins.__import__

    def host_import(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "sys" and (globals or {}).get("__file__", "").endswith("test_macos_word_rules.py"):
            return SimpleNamespace(platform=HOST)
        return real_import(name, globals, locals, fromlist, level)

    builtins.__import__ = host_import

    class NoNativeCompiler:
        def pytest_collection_modifyitems(self, items):
            for item in items:
                if item.module.__name__.endswith("test_macos_word_rules"):
                    item.module.subprocess = SimpleNamespace(run=self.missing_compiler)

        @staticmethod
        def missing_compiler(args, **kwargs):
            raise FileNotFoundError("simulated host has no /usr/bin/osacompile")

    raise SystemExit(pytest.main([
        "-q", "--tb=short",
        "tests/longform_m5/test_public_routing.py",
        "tests/macos_probe/test_longform_evidence.py",
        "tests/test_generation.py",
        "tests/msoffice/test_macos_word_rules.py",
        *sys.argv[2:],
    ], plugins=[NoNativeCompiler()]))
