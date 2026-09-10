"""Portable Office tests must never initialize the real Cocoa automation runtime.

Native acceptance lives in opt-in fixtures outside this suite. Mock runtimes
constructed with __new__ remain available for ABI and process-identity tests.
"""
from __future__ import annotations

from pathlib import Path
import subprocess

import pytest


@pytest.fixture(autouse=True)
def forbid_live_cocoa_runtime(monkeypatch):
    from skills.WPSComposer.scripts.msoffice import macos_excel_launcher, macos_osa_transport

    def forbidden(*args, **kwargs):
        raise AssertionError('Portable Office tests must inject a fake Cocoa runtime')

    monkeypatch.setattr(macos_osa_transport._DarwinRuntime, '__init__', forbidden)

    transport_run = macos_osa_transport.BoundOSAKitTransport.run
    native_transport_helper = Path(macos_osa_transport.__file__).resolve()
    native_runner = subprocess.run

    def guarded_transport_run(self, *args, **kwargs):
        if self._runner is native_runner and self._helper_path == native_transport_helper:
            forbidden()
        return transport_run(self, *args, **kwargs)

    monkeypatch.setattr(
        macos_osa_transport.BoundOSAKitTransport,
        'run',
        guarded_transport_run,
    )

    launcher_init = macos_excel_launcher.BoundedExcelLauncher.__init__

    def guarded_launcher_init(self, *args, **kwargs):
        if not kwargs.get('helper_command'):
            forbidden()
        launcher_init(self, *args, **kwargs)

    monkeypatch.setattr(
        macos_excel_launcher.BoundedExcelLauncher,
        '__init__',
        guarded_launcher_init,
    )
