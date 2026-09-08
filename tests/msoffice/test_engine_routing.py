from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts import orchestrator, conversion, presentation
from skills.WPSComposer.scripts import office_engines as engines


@pytest.mark.parametrize('value', ['word', '', None, 1, True])
def test_invalid_engine_rejected(value):
    with pytest.raises((ValueError, TypeError)):
        orchestrator.generate('# Title', source_is_text=True, engine=value)


def test_explicit_office_rejects_spreadsheet_before_backend(monkeypatch, tmp_path):
    monkeypatch.setattr(orchestrator, 'generate_macos', lambda *a, **k: pytest.fail('launched'))
    with pytest.raises(engines.EngineUnavailableError, match='writer'):
        orchestrator.generate('# Sheet', format='xlsx', source_is_text=True,
                              engine='msoffice', output=str(tmp_path/'out.xlsx'))
    assert not (tmp_path/'out.xlsx').exists()


def test_auto_prefers_wps_and_resolves_only_once(monkeypatch):
    seen=[]
    monkeypatch.setattr(engines, 'engine_executable', lambda engine, component: seen.append(engine) or '/installed/app')
    assert engines.resolve_engine('auto', 'writer') == 'wps'
    assert seen == ['wps']


def test_auto_word_fallback_only_when_wps_missing(monkeypatch):
    monkeypatch.setattr(engines, 'engine_executable', lambda engine, component: '/Word' if engine=='msoffice' else None)
    assert engines.resolve_engine('auto', 'writer') == 'msoffice'
    with pytest.raises(engines.EngineUnavailableError):
        engines.resolve_engine('auto','presentation')


def test_generate_routes_office_through_shared_longform(monkeypatch,tmp_path):
    monkeypatch.setattr(orchestrator.sys,'platform','darwin')
    seen=[]
    target=tmp_path/'out.docx'
    def run(build,format_name,output,timeout,overwrite,*,engine):
        seen.append((engine,format_name,tuple(op.op for op in build.plan.operations)))
        return SimpleNamespace(path=str(target))
    monkeypatch.setattr(orchestrator,'_generate_longform_outcome',run)
    assert orchestrator.generate('# Title\n\n## Chapter\n\nBody',source_is_text=True,output=str(target),engine='msoffice') == str(target)
    assert seen[0][:2] == ('msoffice','docx')
    assert 'writer.ensure_styles' in seen[0][2]


def test_office_legacy_fails_explicitly(monkeypatch,tmp_path):
    with pytest.raises(engines.EngineUnavailableError,match='legacy'):
        orchestrator.generate('---\nlayout_engine: legacy\n---\n# Title',source_is_text=True,engine='msoffice',output=str(tmp_path/'x.docx'))


@pytest.mark.parametrize('timeout',[0,-1,float('inf'),float('nan'),True])
def test_invalid_deadlines_rejected_before_side_effects(timeout,tmp_path):
    with pytest.raises((ValueError,TypeError)):
        orchestrator.generate('# Title',source_is_text=True,timeout=timeout,output=str(tmp_path/'x.docx'))


def test_pinned_wps_com_excludes_office_and_restores_context():
    original=('KWps.Application','Word.Application')
    assert engines.com_progids(original)==original
    with engines.com_engine('wps'):
        assert engines.com_progids(original)==('KWps.Application',)
    assert engines.com_progids(original)==original


def test_word_presentation_uses_named_mac_application(monkeypatch,tmp_path):
    artifact=tmp_path/'x.docx'; artifact.write_bytes(b'test')
    calls=[]
    monkeypatch.setattr(presentation.sys,'platform','darwin')
    monkeypatch.setattr(presentation.subprocess,'run',lambda argv,**kw: calls.append(argv))
    presentation.present_artifact(artifact,engine='msoffice')
    assert calls==[['open','-a','Microsoft Word',str(artifact)]]


def test_conversion_pins_engine_and_deadline(monkeypatch, tmp_path):
    source = tmp_path / 'source.docx'
    source.write_bytes(b'source')
    calls = []
    monkeypatch.setattr(engines, 'engine_executable', lambda engine, component: '/Word' if engine == 'msoffice' else None)
    def select(request):
        calls.append((request.engine, request.timeout))
        return 'test-word', lambda req: req.output
    monkeypatch.setattr(conversion, '_select_backend', select)
    monkeypatch.setattr(conversion, 'validate_pdf', lambda path: None)
    conversion.convert_to_pdf(str(source), engine='auto', timeout=27)
    assert calls == [('msoffice', 27)]


def test_conversion_office_spreadsheet_rejected_before_launch(monkeypatch, tmp_path):
    source = tmp_path / 'source.xlsx'
    source.write_bytes(b'source')
    monkeypatch.setattr(conversion, '_select_backend', lambda req: pytest.fail('backend launched'))
    with pytest.raises(engines.EngineUnavailableError):
        conversion.convert_to_pdf(str(source), engine='msoffice')


def test_wps_docx_presentation_pins_application(monkeypatch, tmp_path):
    target = tmp_path / 'out.docx'
    target.write_bytes(b'artifact')
    calls = []
    monkeypatch.setattr(presentation.sys, 'platform', 'darwin')
    monkeypatch.setattr(presentation.subprocess, 'run', lambda argv, **kw: calls.append(argv))
    orchestrator._return_artifact(target, open_result=True, engine='wps')
    assert calls == [['open', '-a', '/Applications/wpsoffice.app', str(target)]]


def test_mac_auto_ignores_wps_locations_not_supported_by_runtime(monkeypatch):
    monkeypatch.setattr(engines.sys, 'platform', 'darwin')
    monkeypatch.setattr(Path, 'is_dir', lambda p: str(p) in {'/Applications/WPS Office.app', '/Applications/Microsoft Word.app'})
    assert engines.resolve_engine('auto', 'writer') == 'msoffice'


def test_windows_presentation_does_not_wait_or_terminate_interactive_office(monkeypatch, tmp_path):
    target = tmp_path / 'result.docx'
    target.write_bytes(b'artifact')
    monkeypatch.setattr(presentation.sys, 'platform', 'win32')
    monkeypatch.setattr(engines, 'engine_executable', lambda *args: 'WINWORD.EXE')
    calls = []
    monkeypatch.setattr(presentation.subprocess, 'Popen', lambda argv, **kw: calls.append(argv))
    monkeypatch.setattr(presentation.subprocess, 'run', lambda *a, **k: pytest.fail('must not wait and kill interactive Office'))
    presentation.present_artifact(target, engine='msoffice')
    assert calls == [['WINWORD.EXE', str(target)]]
