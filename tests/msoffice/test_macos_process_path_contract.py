"""Darwin process identities retain POSIX semantics on portable test hosts."""
from pathlib import PureWindowsPath

import pytest

from skills.WPSComposer.scripts.msoffice import macos_osa_transport as osa


def test_mac_process_identity_roundtrips_when_host_paths_are_windows(monkeypatch):
    monkeypatch.setattr(osa, 'Path', PureWindowsPath)
    value = osa.ExcelProcessIdentity(42, 10, 20,
        '/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel',
        'com.microsoft.Excel')
    import json
    request = osa.HelperRequest.from_json(json.dumps(value.to_request(shared_deadline=100)))
    assert request.identity == value
    osa.require_excel_process(value)


@pytest.mark.parametrize('executable', ['', 'Applications/Microsoft Excel', r'C:\Office\Microsoft Excel', '../Microsoft Excel'])
def test_mac_process_identity_rejects_non_posix_absolute_path(executable):
    with pytest.raises(ValueError, match='absolute'):
        osa.ExcelProcessIdentity(42, 10, 20, executable, 'com.microsoft.Excel')


def test_darwin_router_rejects_windows_backslash_aliases(monkeypatch):
    import ntpath
    from types import SimpleNamespace
    from dataclasses import replace
    monkeypatch.setattr(osa, 'os', SimpleNamespace(path=ntpath))
    monkeypatch.setattr(osa, 'Path', PureWindowsPath)
    expected = osa.ExcelProcessIdentity(42, 10, 20,
        '/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel', 'com.microsoft.Excel')
    osa.require_same_process(expected, replace(expected))
    # Backslash is a literal Darwin filename character, not a separator.
    altered = replace(expected, executable='/Applications/Microsoft Excel.app/Contents/MacOS\\Microsoft Excel')
    with pytest.raises(osa.OSATransportError, match='EXECUTABLE_CHANGED'):
        osa.require_same_process(expected, altered)
    router = osa.BoundSendRouter(expected, SimpleNamespace(snapshot=lambda pid: altered), helper_pid=77)
    assert router._is_verified_excel_target(osa.AppleEventTarget.application_path(expected.executable))
    assert router._is_verified_excel_target(osa.AppleEventTarget.application_path('/Applications/Microsoft Excel.app'))
    assert not router._is_verified_excel_target(osa.AppleEventTarget.application_path(altered.executable))
    assert not router._is_verified_excel_target(osa.AppleEventTarget.kernel_pid(42))
