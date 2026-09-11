from types import SimpleNamespace
import time


def test_runtime_retains_native_execution_issues_for_later_quality_passes(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts.msoffice import macos_runtime as runtime
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    build = build_longform_generation('# Title\n\nBody.')
    adapter = runtime.MacWordAdapter(build)
    adapter.staging_root = tmp_path
    issue = SimpleNamespace(code='EQUATION_INSERT_FAILED')
    outcome = SimpleNamespace(issues=(issue,))
    compiled = SimpleNamespace(source='synthetic', nodes={}, operations=1, issues=())
    monkeypatch.setattr(runtime, 'compile_plan', lambda *a, **kw: compiled)
    monkeypatch.setattr(runtime, '_build_executor_resources', lambda *a: ())
    monkeypatch.setattr(adapter, '_ensure_started', lambda *a: None)
    monkeypatch.setattr(adapter, '_run', lambda *a: 'synthetic native fallback report')
    monkeypatch.setattr(runtime, 'parse_result', lambda *a: outcome)
    monkeypatch.setattr(runtime, 'validate_office_package', lambda *a: None)
    assert adapter.execute(build, (), time.monotonic()+30) is outcome
    assert adapter.issues == (issue,)
