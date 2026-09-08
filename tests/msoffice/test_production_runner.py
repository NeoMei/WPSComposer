import importlib.util
import json
from pathlib import Path
import zipfile

import pytest

_SPEC = importlib.util.spec_from_file_location('production_runner', Path(__file__).resolve().parents[2] / 'fixtures' / 'verify_msoffice_production.py')
runner = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(runner)


def document_package(path, *, numbering=True, body_font='仿宋', indent='480'):
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    p = '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>正文验收文本。</w:t></w:r></w:p>'
    h = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>章节验收</w:t></w:r></w:p>'
    styles = f'<w:styles xmlns:w="{ns}"><w:style w:styleId="BodyText"><w:rPr><w:rFonts w:eastAsia="{body_font}"/><w:sz w:val="24"/></w:rPr><w:pPr><w:ind w:firstLine="{indent}"/></w:pPr></w:style><w:style w:styleId="Heading1"><w:pPr><w:outlineLvl w:val="0"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style></w:styles>'
    num = f'<w:numbering xmlns:w="{ns}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:pStyle w:val="Heading1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
    with zipfile.ZipFile(path, 'w') as package:
        package.writestr('word/document.xml', f'<w:document xmlns:w="{ns}"><w:body>{h}{p}</w:body></w:document>')
        package.writestr('word/styles.xml', styles)
        if numbering:
            package.writestr('word/numbering.xml', num)


def test_docx_checks_effective_style_body_and_native_numbering(tmp_path):
    docx = tmp_path / 'native.docx'
    document_package(docx)
    spec = {'title': None, 'body': '正文验收文本。', 'headings': ['章节验收'], 'table_rows': [], 'toc': False, 'numbered': True, 'markers': []}
    result = runner.inspect_docx(docx, spec)
    assert result['checks']['body_fangsong_12pt_two_character_indent']
    assert result['checks']['heading_sizes_and_outline_levels']
    assert result['checks']['native_heading_numbering']


@pytest.mark.parametrize('change', [{'numbering': False}, {'body_font': 'Arial'}, {'indent': '0'}])
def test_docx_guardrails_reject_native_regressions(tmp_path, change):
    docx = tmp_path / 'broken.docx'
    document_package(docx, **change)
    spec = {'title': None, 'body': '正文验收文本。', 'headings': ['章节验收'], 'table_rows': [], 'toc': False, 'numbered': True, 'markers': []}
    result = runner.inspect_docx(docx, spec)
    assert not all(result['checks'].values())


def test_failed_public_call_preserves_structured_report_and_actual_source_digest(monkeypatch, tmp_path):
    calls = []

    def generation(source, **kwargs):
        calls.append((source, kwargs))
        raise RuntimeError('native host unavailable')

    import skills.WPSComposer as public
    monkeypatch.setattr(public, 'generate', generation)
    root = tmp_path / 'acceptance'
    result = runner.run_acceptance(root, fixture='smoke', engine='msoffice', timeout=30)
    report = json.loads((root / 'report.json').read_text())
    assert result['status'] == report['status'] == 'FAIL'
    assert report['errors'][0]['type'] == 'RuntimeError'
    assert report['source']['sha256'] == runner.sha256(root / 'source.md')
    assert report['candidate']['source_files']
    assert calls[0][1]['engine'] == 'msoffice'
    assert calls[0][1]['format'] == 'docx'
    assert 0 < calls[0][1]['timeout'] <= 30
    assert calls[0][1]['open_result'] is False
    assert report['native_ui_acceptance'] == 'NOT_RUN'


def test_existing_output_root_is_never_reused(tmp_path):
    root = tmp_path / 'old-evidence'
    root.mkdir()
    (root / 'report.json').write_text('{"status":"PASS"}')
    with pytest.raises(FileExistsError):
        runner.run_acceptance(root)
    assert json.loads((root / 'report.json').read_text())['status'] == 'PASS'


@pytest.mark.parametrize('timeout', [0, -1, float('inf'), float('nan'), True])
def test_invalid_total_budget_does_not_create_evidence_directory(tmp_path, timeout):
    root = tmp_path / 'invalid'
    with pytest.raises(ValueError):
        runner.run_acceptance(root, timeout=timeout)
    assert not root.exists()


def test_fixture_contract_matches_actual_shared_plan_heading_and_numbering_depth():
    from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
    for name in ('smoke', 'representative'):
        content, spec = runner.fixture_content(name)
        build = build_longform_generation(content)
        headings = [op for op in build.plan.operations if op.op == 'writer.add_heading']
        assert [op.args['text'] for op in headings] == spec['headings']
        if spec['numbered']:
            assert sum(op.args.get('numbering') is True for op in headings) == spec['numbered_levels']


def test_explicit_zero_character_indent_overrides_stale_point_indent(tmp_path):
    docx = tmp_path / 'broken.docx'
    document_package(docx)
    with zipfile.ZipFile(docx) as package:
        members = {name: package.read(name) for name in package.namelist()}
    members['word/styles.xml'] = members['word/styles.xml'].replace(b'w:firstLine="480"', b'w:firstLine="480" w:firstLineChars="0"')
    with zipfile.ZipFile(docx, 'w') as package:
        for name, payload in members.items():
            package.writestr(name, payload)
    spec = {'title': None, 'body': '正文验收文本。', 'headings': ['章节验收'], 'table_rows': [], 'toc': False, 'numbered': True, 'markers': []}
    assert not runner.inspect_docx(docx, spec)['checks']['body_fangsong_12pt_two_character_indent']


def test_pdf_inspector_checks_real_pdf_text_and_nonempty_pages(tmp_path):
    reportlab = pytest.importorskip('reportlab.pdfgen.canvas')
    pdf = tmp_path / 'native.pdf'
    canvas = reportlab.Canvas(str(pdf))
    canvas.drawString(50, 750, 'Heading acceptance')
    canvas.drawString(50, 730, 'Body acceptance')
    canvas.drawString(50, 710, 'END-SENTINEL')
    canvas.save()
    spec = {'title': None, 'body': 'Body acceptance', 'headings': ['Heading acceptance'], 'table_rows': [], 'markers': ['END-SENTINEL']}
    assert all(runner.inspect_pdf(pdf, spec)['checks'].values())
    spec['body'] = 'Missing native content'
    assert not runner.inspect_pdf(pdf, spec)['checks']['pdf_required_content']
