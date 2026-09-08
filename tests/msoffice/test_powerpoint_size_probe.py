"""Preparation checks only: these tests never execute AppleEvents."""
from __future__ import annotations

import importlib
import math
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile

import pytest


def probe():
    return importlib.import_module('fixtures.microsoft_parity.macos_powerpoint_size_probe')


def test_requested_size_compiles_opposite_orientation_width_swap_width():
    steps = probe().compile_steps(720, 405)
    assert [step['command'] for step in steps] == [
        'set slide orientation of page setup of ownedDoc to vertical orientation',
        'set slide width of page setup of ownedDoc to 405',
        'set slide orientation of page setup of ownedDoc to horizontal orientation',
        'set slide width of page setup of ownedDoc to 720',
    ]
    tall = probe().compile_steps(405, 720)
    assert 'horizontal orientation' in tall[0]['command']
    assert 'vertical orientation' in tall[2]['command']
    assert all('slide height' not in step['command'] for step in steps)


@pytest.mark.parametrize('width,height', [(True, 400), (0, 400), (-1, 400),
    (math.inf, 400), (500, math.nan), ('720', 405), (720, 10**400)])
def test_invalid_size_is_rejected_before_any_native_work(width, height):
    with pytest.raises(ValueError):
        probe().compile_steps(width, height)


def test_import_is_inert_and_module_compiles(tmp_path):
    script = '''
import subprocess
def forbidden(*a, **kw): raise AssertionError('native process at import')
subprocess.run = forbidden
import fixtures.microsoft_parity.macos_powerpoint_size_probe
'''
    subprocess.run([sys.executable, '-c', script], check=True,
                   cwd=Path(__file__).resolve().parents[2], timeout=10)
    source = Path(probe().__file__).read_text()
    compile(source, probe().__file__, 'exec')


def test_execution_requires_explicit_guard_and_platform_before_output(tmp_path, monkeypatch):
    module = probe()
    destination = tmp_path / 'never-created'
    with pytest.raises(ValueError, match='execute'):
        module.run(destination, [(720, 405)], sentinel_names=['sentinel'])
    monkeypatch.setattr(module.sys, 'platform', 'linux')
    with pytest.raises(RuntimeError, match='macOS'):
        module.run(destination, [(720, 405)], sentinel_names=['sentinel'], execute=True)
    assert not destination.exists()


def test_owned_guard_rejects_attached_external_or_unentered_session(tmp_path):
    from types import SimpleNamespace
    root = tmp_path / 'container'
    job = root / 'session-owned'
    good = dict(_entered=True, _attached=False, _uncertain=False,
                _job=job, _path=str(job / 'bound-owned.pptx'))
    probe().require_owned(SimpleNamespace(**good), root)
    for changed in ({'_attached': True}, {'_entered': False}, {'_uncertain': True},
                    {'_path': str(tmp_path / 'user.pptx')}, {'_job': root / 'external'}):
        with pytest.raises(RuntimeError):
            probe().require_owned(SimpleNamespace(**{**good, **changed}), root)


def test_read_size_is_read_only_ooxml_diagnostic(tmp_path):
    path = tmp_path / 'sample.pptx'
    with ZipFile(path, 'w') as package:
        package.writestr('ppt/presentation.xml',
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:sldSz cx="9144000" cy="5143500"/></p:presentation>')
    before = path.read_bytes()
    assert probe().read_size(path) == {'width_pt': 720.0, 'height_pt': 405.0}
    assert path.read_bytes() == before


def test_native_failure_retains_owned_document_without_any_later_call():
    from types import SimpleNamespace
    calls = []
    session = SimpleNamespace(_entered=True, _uncertain=False,
        _quarantine=lambda reason: calls.append(('quarantine', reason)),
        _release=lambda: calls.append(('release',)),
        close=lambda **kw: calls.append(('close',)))
    probe().retain_failure(session, RuntimeError('uncertain'))
    assert [call[0] for call in calls] == ['quarantine', 'release']
    assert not session._entered


def test_default_cli_prepares_only_even_when_run_would_be_available(monkeypatch, capsys):
    module = probe()
    monkeypatch.setattr(module, 'run', lambda *a, **kw: pytest.fail('native execution requested'))
    assert module.main([]) == 0
    assert 'PREPARED_NOT_EXECUTED' in capsys.readouterr().out


def test_each_native_session_gets_at_most_sixty_seconds(monkeypatch):
    from types import SimpleNamespace
    module = probe()
    monkeypatch.setattr(module.time, 'monotonic', lambda: 100)
    session = SimpleNamespace(__enter__=lambda: None)
    module._deadline_session(session, 700)
    assert session.timeout == 60


@pytest.mark.parametrize('timeout', [False, True])
def test_inventory_never_persists_unrelated_text_even_partial_stdout(tmp_path, monkeypatch, timeout):
    import json
    from types import SimpleNamespace
    from skills.WPSComposer.scripts.msoffice import macos_office_runtime
    module = probe()
    marker = 'PRIVATE-UNRELATED-PRESENTATION-CONTENT-8492'
    class Lock:
        def __init__(self, *args): pass
        def acquire(self, *args): pass
        def close(self): pass
        def quarantine(self, *args): pass
    monkeypatch.setattr(macos_office_runtime, 'OfficeJobLock', Lock)
    def execute(*args, **kwargs):
        stdout = json.dumps([['sentinel', '/synthetic.pptx', False, 1, [marker], 720, 'horizontal']])
        if timeout:
            raise subprocess.TimeoutExpired('osascript', 25, output=stdout.encode(), stderr=b'timed out')
        return SimpleNamespace(returncode=0, stdout=stdout, stderr='')
    monkeypatch.setattr(module.subprocess, 'run', execute)
    output = tmp_path / 'before.json'
    if timeout:
        with pytest.raises(subprocess.TimeoutExpired):
            module._inventory(output, tmp_path / 'container', module.time.monotonic() + 60)
    else:
        before = module._inventory(output, tmp_path / 'container', module.time.monotonic() + 60)
        marker += '-changed'
        after = module._inventory(tmp_path / 'after.json', tmp_path / 'container', module.time.monotonic() + 60)
        assert before != after
        assert before[0][4]['has_text'] is True
    for path in tmp_path.rglob('*'):
        if path.is_file():
            assert b'PRIVATE-UNRELATED-PRESENTATION-CONTENT-8492' not in path.read_bytes()


def test_probe_uses_real_blank_slide_return_contract(monkeypatch):
    from skills.WPSComposer.scripts.msoffice.macos_powerpoint_session import MacPowerPointSession
    session = MacPowerPointSession()
    monkeypatch.setattr(session, '_semantic_slide', lambda *args: 1)
    captured = []
    def textbox(index, *args):
        assert type(index) is int
        captured.append(index)
    monkeypatch.setattr(session, 'add_textbox', textbox)
    probe()._add_probe_slide(session, 720, 405)
    assert captured == [1]
