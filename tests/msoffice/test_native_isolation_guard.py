from __future__ import annotations

import sys
import time

import pytest


def _identity():
    from skills.WPSComposer.scripts.msoffice.macos_osa_transport import ExcelProcessIdentity

    return ExcelProcessIdentity(
        4242,
        100,
        250,
        "/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel",
        "com.microsoft.Excel",
    )


def test_global_guard_rejects_default_osa_helper_before_deadline_logic(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_osa_transport import BoundOSAKitTransport

    script = tmp_path / "never-run.applescript"
    script.write_text('return "never run"')
    transport = BoundOSAKitTransport(_identity())

    with pytest.raises(AssertionError, match="Portable Office tests"):
        transport.run(script, deadline=time.monotonic() - 1)


def test_global_guard_rejects_default_excel_launcher_before_spawn(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_excel_launcher import BoundedExcelLauncher

    with pytest.raises(AssertionError, match="Portable Office tests"):
        BoundedExcelLauncher(
            evidence_dir=tmp_path / "launcher",
            deadline=time.monotonic() + 5,
        )


@pytest.mark.skipif(
    sys.platform == "win32",
    reason="Bounded launcher uses POSIX inherited file descriptors",
)
def test_global_guard_allows_explicit_fake_helper_and_new_only_runtime(tmp_path):
    from skills.WPSComposer.scripts.msoffice.macos_excel_launcher import BoundedExcelLauncher
    from skills.WPSComposer.scripts.msoffice.macos_osa_transport import _DarwinRuntime

    runtime = _DarwinRuntime.__new__(_DarwinRuntime)
    assert isinstance(runtime, _DarwinRuntime)

    launcher = BoundedExcelLauncher(
        evidence_dir=tmp_path / "launcher",
        deadline=time.monotonic() + 5,
        helper_command=[sys.executable, "-c", "raise SystemExit(0)"],
    )
    launcher.abort()
