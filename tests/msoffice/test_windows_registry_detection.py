"""Read-only discovery of out-of-process COM servers across registry views."""
from contextlib import contextmanager
from types import SimpleNamespace
import sys

import pytest

from skills.WPSComposer.scripts import office_engines as engines


@pytest.fixture
def registry(monkeypatch):
    values, reads = {}, []
    @contextmanager
    def open_key(root, path, reserved=0, access=1):
        assert root == 'HKCR'
        assert access & 1, 'Only KEY_READ access is allowed'
        view = access & (256 | 512)
        reads.append((view, path))
        if (view, path) not in values:
            raise FileNotFoundError(path)
        yield values[view, path]
    monkeypatch.setitem(sys.modules, 'winreg', SimpleNamespace(
        HKEY_CLASSES_ROOT='HKCR', KEY_READ=1, KEY_WOW64_32KEY=512,
        KEY_WOW64_64KEY=256, OpenKey=open_key,
        QueryValueEx=lambda key, name: (key, 1)))
    monkeypatch.setattr(engines.sys, 'platform', 'win32')
    return values, reads


def register(values, view, progid, executable):
    clsid = '{' + progid + '}'
    values[view, progid + '\\CLSID'] = clsid
    values[view, 'CLSID\\' + clsid + '\\LocalServer32'] = f'"{executable}" /Automation'


@pytest.mark.parametrize('view', [512, 256])
def test_auto_discovers_wps_in_alternate_registry_view(registry, tmp_path, view):
    values, reads = registry
    executable = tmp_path / 'wps.exe'; executable.touch()
    register(values, view, 'KWps.Application', executable)
    assert engines.engine_executable('wps') == str(executable)
    assert engines.resolve_engine('auto', 'writer') == 'wps'
    assert any(read_view == view for read_view, path in reads)


def test_default_registry_view_keeps_precedence(registry, tmp_path):
    values, _ = registry
    paths = []
    for view in (0, 512, 256):
        folder = tmp_path / str(view); folder.mkdir()
        path = folder / 'wps.exe'; path.touch(); paths.append(path)
        register(values, view, 'KWps.Application', path)
    assert engines.engine_executable('wps') == str(paths[0])


def test_missing_registered_file_does_not_hide_valid_alternate_view(registry, tmp_path):
    values, _ = registry
    register(values, 0, 'KWps.Application', tmp_path / 'missing' / 'wps.exe')
    path = tmp_path / 'wps.exe'; path.touch()
    register(values, 512, 'KWps.Application', path)
    assert engines.engine_executable('wps') == str(path)


def test_registry_views_cannot_be_combined_into_false_registration(registry, tmp_path):
    values, _ = registry
    path = tmp_path / 'wps.exe'; path.touch()
    values[0, 'KWps.Application\\CLSID'] = '{WPS}'
    values[512, 'CLSID\\{WPS}\\LocalServer32'] = f'"{path}"'
    assert engines.engine_executable('wps') is None
