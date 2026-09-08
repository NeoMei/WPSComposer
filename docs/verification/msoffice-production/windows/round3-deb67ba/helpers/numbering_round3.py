"""Supplemental real Word cases using production composer/executor and public generate."""
import sys, json, traceback, hashlib, re, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from PIL import Image
import fitz
from skills.WPSComposer import generate
from skills.WPSComposer.scripts.msoffice.windows_host import create_dedicated_composer, _word_processes
from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.generation_plan import GenerationOperation

root = Path('build/msoffice-production/numbering-03').resolve()
root.mkdir(exist_ok=False)
report = {'status': 'RUNNING', 'checks': {}, 'errors': []}
def persist():
    (root/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
composer = None
try:
    composer = create_dedicated_composer(str(root))
    report['ordinary_host'] = {'pid': composer.identity.pid, 'binding': composer.application_binding}
    composer.reset()
    composer.ensure_heading_styles({1: {'fontName': '黑体', 'fontSize': 16, 'bold': True}})
    executor = WindowsLongformExecutor(composer_factory=lambda: composer)
    ops = [
        GenerationOperation(op='writer.add_heading', args={'text': 'Ordinary unnumbered heading', 'level': 1, 'numbering': False}),
        GenerationOperation(op='writer.add_heading', args={'text': 'Numbered chapter', 'level': 1, 'numbering': True, 'numberingScheme': 'decimal'}),
    ]
    for op in ops:
        executor._run_op(composer, op)
    composer.update_fields()
    paragraphs = []
    for i in range(1, composer.doc.Paragraphs.Count + 1):
        p = composer.doc.Paragraphs(i)
        if p.Range.Text.strip():
            paragraphs.append({'text': p.Range.Text.strip(), 'list_string': str(p.Range.ListFormat.ListString), 'outline_level': int(p.OutlineLevel), 'style': str(p.Style.NameLocal), 'font_size': float(p.Range.Font.Size)})
    report['ordinary_paragraphs'] = paragraphs
    first, second = paragraphs[:2]
    report['checks']['ordinary_h1_stays_unnumbered_after_numbered_h1'] = first['list_string'] == '' and first['outline_level'] == 1 and second['list_string'] == '1'
    composer.save_docx(root/'ordinary.docx')
    composer.export_pdf(root/'ordinary.pdf')
    composer.close(save_changes=False)
    report['ordinary_host']['closed'] = composer._closed
    report['ordinary_host']['pid_absent'] = composer.identity.pid not in _word_processes()
    composer = None
    persist()

    # Simple synthetic fixture, independent of user images.
    Image.new('RGB', (400, 160), (50, 100, 180)).save(root/'figure.png')
    content = '''---
title: Native transparent heading
title_page: false
toc: true
heading_numbering: decimal
caption_numbering: auto
---
# 1 Numbered chapter

:::figure {#fig:first caption="First numbered figure"}
![first](figure.png)
:::

# 第二章 Unnumbered boundary

:::figure {#fig:second caption="Second numbered figure"}
![second](figure.png)
:::

:::table {#tab:values caption="Measured values"}
| Signal | Value |
| --- | --- |
| Alpha | 10 |
| Beta | 20 |
:::

Reference pair: {{ref:fig:first}} and {{ref:fig:second}}. Table reference: {{ref:tab:values}}.
'''
    source = root/'transparent.md'
    source.write_text(content, encoding='utf-8')
    build = build_longform_generation(content, base_dir=str(root))
    (root/'transparent-plan.json').write_text(json.dumps(build.plan.to_dict(), ensure_ascii=False, indent=2), encoding='utf-8')
    headings = [dict(op.args) for op in build.plan.operations if op.op == 'writer.add_heading']
    report['transparent_heading_operations'] = headings
    report['checks']['shared_plan_marks_boundary_sequence_transparent'] = any(h.get('sequenceTransparent') is True and h.get('numbering') is False for h in headings)
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    start = time.monotonic()
    result = generate(str(source), format='docx', output=str(root/'transparent.docx'), engine='msoffice', timeout=600, open_result=False)
    report['public_generate_seconds'] = round(time.monotonic() - start, 3)
    from skills.WPSComposer import convert_to_pdf
    convert_to_pdf(result, str(root/'transparent.pdf'), engine='msoffice', timeout=600, open_result=False)
    with fitz.open(root/'transparent.pdf') as pdf:
        text = '\n'.join(p.get_text() for p in pdf)
        report['transparent_pdf_pages'] = len(pdf)
    report['transparent_pdf_text'] = text
    compact = re.sub(r'\s+', '', text)
    report['checks']['chapter_figures_continue_1_1_then_1_2'] = '图1-1Firstnumberedfigure' in compact and '图1-2Secondnumberedfigure' in compact
    report['checks']['native_figure_and_table_references'] = 'Referencepair:图1-1and图1-2.' in compact and 'Tablereference:表1-1.' in compact
    report['checks']['table_caption_and_values'] = '表1-1Measuredvalues' in compact and 'Alpha10' in compact and 'Beta20' in compact
    import zipfile, xml.etree.ElementTree as ET
    with zipfile.ZipFile(root/'transparent.docx') as package:
        xml=ET.fromstring(package.read('word/document.xml'))
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    fields=[n.text or '' for n in xml.findall('.//w:instrText',ns)]
    report['native_field_codes']=fields
    report['checks']['two_native_images_and_three_native_refs'] = len(xml.findall('.//w:drawing',ns))==2 and sum(bool(re.match(r'\s*REF\b',f)) for f in fields)==3
    report['checks']['source_unchanged'] = hashlib.sha256(source.read_bytes()).hexdigest() == before
    report['status'] = 'PASS' if all(report['checks'].values()) else 'FAIL'
except Exception as exc:
    report['status'] = 'FAIL'
    report['errors'].append({'error': repr(exc), 'traceback': traceback.format_exc()})
finally:
    if composer is not None:
        try:
            composer.close(save_changes=False)
            report['ordinary_host']['closed'] = composer._closed
        except Exception as exc:
            report['errors'].append({'cleanup_error': repr(exc)})
    persist()
print(json.dumps(report, ensure_ascii=False))
raise SystemExit(0 if report['status'] == 'PASS' else 1)
