from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts import orchestrator, conversion, presentation
from skills.WPSComposer.scripts import office_engines as engines


@pytest.mark.parametrize('value', ['word', '', None, 1, True])
def test_invalid_engine_rejected(value):
    with pytest.raises((ValueError, TypeError)):
        orchestrator.generate('# Title', source_is_text=True, engine=value)


@pytest.mark.parametrize('fmt', ['xlsx', 'pptx'])
def test_explicit_office_routes_to_native_component(monkeypatch, tmp_path, fmt):
    from skills.WPSComposer.scripts.msoffice import macos_office_runtime
    monkeypatch.setattr(orchestrator.sys, 'platform', 'darwin')
    monkeypatch.setattr(orchestrator, 'generate_macos', lambda *a, **k: pytest.fail('WPS launched'))
    calls = []
    def run(doc, format_name, output, preset, **kwargs):
        calls.append((format_name, kwargs['timeout']))
        return output
    monkeypatch.setattr(macos_office_runtime, 'generate', run)
    out = tmp_path / ('out.' + fmt)
    assert orchestrator.generate('# Sheet', format=fmt, source_is_text=True,
                                engine='msoffice', output=str(out), timeout=27) == str(out)
    assert calls == [(fmt, 27)]


def test_auto_prefers_wps_and_resolves_only_once(monkeypatch):
    seen=[]
    monkeypatch.setattr(engines, 'engine_executable', lambda engine, component: seen.append(engine) or '/installed/app')
    assert engines.resolve_engine('auto', 'writer') == 'wps'
    assert seen == ['wps']


def test_auto_word_fallback_only_when_wps_missing(monkeypatch):
    monkeypatch.setattr(engines, 'engine_executable', lambda engine, component: '/Word' if engine=='msoffice' else None)
    assert engines.resolve_engine('auto', 'writer') == 'msoffice'
    assert engines.resolve_engine('auto', 'presentation') == 'msoffice'


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


@pytest.mark.parametrize('suffix,component', [('xlsx','spreadsheet'),('pptx','presentation')])
def test_conversion_office_component_is_pinned(monkeypatch, tmp_path, suffix, component):
    from skills.WPSComposer.scripts.msoffice import macos_office_runtime
    source = tmp_path / ('source.' + suffix)
    source.write_bytes(b'synthetic source')
    monkeypatch.setattr(conversion.sys, 'platform', 'darwin')
    calls = []
    def convert(req, **kwargs):
        calls.append((req.component, req.engine, kwargs['timeout']))
        return req.output
    monkeypatch.setattr(macos_office_runtime, 'convert', convert)
    monkeypatch.setattr(conversion, 'validate_pdf', lambda path: None)
    conversion.convert_to_pdf(str(source), engine='msoffice', timeout=42)
    assert calls == [(component, 'msoffice', 42)]


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
    monkeypatch.setattr(Path, 'is_dir', lambda p: p in {Path('/Applications/WPS Office.app'), Path('/Applications/Microsoft Word.app')})
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


@pytest.mark.parametrize('component,progid', [('writer','Word.Application'), ('spreadsheet','Excel.Application'), ('presentation','PowerPoint.Application')])
def test_pinned_microsoft_com_never_falls_back_to_wps(component, progid):
    values = ('KWps.Application', 'Ket.Application', 'Wpp.Application', progid)
    with engines.com_engine('msoffice'):
        assert engines.com_progids(values) == (progid,)
        with pytest.raises(engines.EngineUnavailableError):
            engines.com_progids(('Ket.Application',))
    assert engines.com_progids(values) == values


@pytest.mark.parametrize('suffix,app', [('xlsx','Microsoft Excel'), ('pptx','Microsoft PowerPoint')])
def test_office_open_result_uses_selected_app(monkeypatch, tmp_path, suffix, app):
    target = tmp_path / ('result.' + suffix)
    target.write_bytes(b'synthetic output')
    monkeypatch.setattr(presentation.sys, 'platform', 'darwin')
    calls = []
    monkeypatch.setattr(presentation.subprocess, 'run', lambda argv, **kw: calls.append(argv))
    orchestrator._return_artifact(target, open_result=True, engine='msoffice')
    assert calls == [['open','-a',app,str(target)]]
