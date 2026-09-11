from __future__ import annotations

from pathlib import Path
import subprocess


def test_locked_qs_security_and_express_request_compatibility():
    """Exercise both actual consumers after the repository's required npm ci."""
    result = subprocess.run(
        ["node", "--test", str(Path(__file__).with_name("qs_dependency_regressions.cjs"))],
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 0, result.stdout + result.stderr
