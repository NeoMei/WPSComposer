#!/usr/bin/env python3
"""Run actual public DOCX generation/PDF conversion in a fresh evidence directory.

Example (Windows or macOS, from the candidate checkout):
    python fixtures/verify_msoffice_production.py --output-root build/word-smoke --fixture smoke --timeout 180
    python fixtures/verify_msoffice_production.py --output-root build/word-representative --timeout 600
Use --engine wps for the separate WPS regression gate. PASS covers these native
API and artifact checks only; real UI edit/undo/save/reopen remains NOT_RUN.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import platform
import posixpath
import re
import subprocess
import sys
import time
import traceback
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{' + NS['w'] + '}'
BODY = '这是原生排版正文验收段落，用于检查仿宋十二磅和两字符首行缩进，同时确保中文内容完整保留。'
HEADINGS = ['一级章节验收', '二级章节验收', '三级章节验收', '四级章节验收', '五级章节验收', '六级章节验收']
END = 'NATIVE-PRODUCTION-END'


def fixture_content(name):
    if name == 'smoke':
        text = '---\ntitle: 原生烟雾验收报告\ntoc: false\ntitle_page: false\nheading_numbering: none\n---\n# 原生烟雾验收\n\n' + BODY + '\n\n' + END + '\n'
        return text, {'title': None, 'body': BODY, 'headings': ['原生烟雾验收'], 'table_rows': [], 'toc': False, 'numbered': False, 'markers': [END]}
    if name != 'representative':
        raise ValueError('Unknown acceptance fixture')
    title = '原生文档生产验收报告'
    rows = [['序号', '验收项目', '结果']] + [[str(i), f'长表第{i:02d}项', '内容完整'] for i in range(1, 37)]
    text = f'---\ntitle: {title}\nauthor: WPSComposer acceptance\ndate: 2026-09-08\ntoc: true\ntitle_page: true\nheading_numbering: decimal\n---\n'
    text += '\n# ' + HEADINGS[0] + '\n\n' + BODY + '\n\n'
    for level, heading in enumerate(HEADINGS[1:], 2):
        text += '#' * level + ' ' + heading + '\n\n' + f'层级{level}内容保留。' + '\n\n'
    text += '| ' + ' | '.join(rows[0]) + ' |\n| --- | --- | --- |\n'
    text += ''.join('| ' + ' | '.join(row) + ' |\n' for row in rows[1:])
    text += '\n' + END + '\n'
    return text, {'title': title, 'body': BODY, 'headings': HEADINGS, 'table_rows': rows, 'toc': True, 'numbered': True, 'numbered_levels': 4, 'markers': [END], 'section_page_number_policy': True}


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def _text(node):
    return ''.join(item.text or '' for item in node.findall('.//w:t', NS))


def _value(node, path, default=None):
    item = node.find(path, NS)
    return item.get(W + 'val', default) if item is not None else default


def _style_chain(paragraph, styles):
    chain = []
    seen = set()
    style_id = _value(paragraph, 'w:pPr/w:pStyle')
    while style_id:
        if style_id in seen or style_id not in styles:
            raise ValueError('Missing or cyclic native paragraph style')
        seen.add(style_id)
        style = styles[style_id]
        chain.append(style)
        style_id = _value(style, 'w:basedOn')
    return list(reversed(chain)) + [paragraph]


def _effective(paragraph, styles, path, attribute='val'):
    value = None
    for node in _style_chain(paragraph, styles):
        item = node.find(path, NS)
        if item is not None and W + attribute in item.attrib:
            value = item.get(W + attribute)
    return value


def _run_property(paragraph, styles, name, attribute='val'):
    inherited = _effective(paragraph, styles, 'w:rPr/w:' + name, attribute)
    values = []
    for run in paragraph.findall('.//w:r', NS):
        if not _text(run).strip():
            continue
        item = run.find('w:rPr/w:' + name, NS)
        values.append(item.get(W + attribute, inherited) if item is not None else inherited)
    return values


def _heading_numbered(paragraph, styles, numbering, level):
    if numbering is None:
        return False
    num_id = _effective(paragraph, styles, 'w:pPr/w:numPr/w:numId')
    if num_id in (None, '0'):
        return False
    instance = next((x for x in numbering.findall('w:num', NS) if x.get(W + 'numId') == num_id), None)
    if instance is None:
        return False
    abstract_id = _value(instance, 'w:abstractNumId')
    abstract = next((x for x in numbering.findall('w:abstractNum', NS) if x.get(W + 'abstractNumId') == abstract_id), None)
    if abstract is None:
        return False
    native_level = next((x for x in abstract.findall('w:lvl', NS) if x.get(W + 'ilvl') == str(level)), None)
    if native_level is None or _value(native_level, 'w:numFmt') != 'decimal':
        return False
    # Word can bind ilvl through the style-linked abstract level instead of
    # placing it directly on the paragraph/style numPr.
    ilvl = _effective(paragraph, styles, 'w:pPr/w:numPr/w:ilvl')
    if ilvl is not None and ilvl != str(level):
        return False
    if ilvl is None and _value(native_level, 'w:pStyle') != _value(paragraph, 'w:pPr/w:pStyle'):
        return False
    return '%' + str(level + 1) in (_value(native_level, 'w:lvlText', '') or '')


def _section_numbering_policy(package, document):
    sections = document.findall('.//w:sectPr', NS)
    if len(sections) != 3:
        return False
    for section, expected_format in zip(sections[1:], ['lowerRoman', 'decimal']):
        number = section.find('w:pgNumType', NS)
        if number is None or number.get(W + 'start') != '1' or number.get(W + 'fmt', 'decimal') != expected_format:
            return False
    relpath = 'word/_rels/document.xml.rels'
    if relpath not in package.namelist():
        return False
    rels = ET.fromstring(package.read(relpath))
    targets = {r.get('Id'): r.get('Target') for r in rels if r.get('TargetMode') != 'External'}
    rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    prior = {}
    for index, section in enumerate(sections):
        for ref in section.findall('w:footerReference', NS):
            target = targets.get(ref.get(rid))
            if not target:
                return False
            path = posixpath.normpath(posixpath.join('word', target)) if not target.startswith('/') else target.lstrip('/')
            if not path.startswith('word/') or path not in package.namelist():
                return False
            prior[ref.get(W + 'type', 'default')] = ET.fromstring(package.read(path))
        has_page = any(
            re.search(r'\bPAGE\b', text)
            for footer in prior.values()
            for text in ([n.text or '' for n in footer.findall('.//w:instrText', NS)] + [n.get(W + 'instr', '') for n in footer.findall('.//w:fldSimple', NS)])
        )
        if bool(has_page) != (index > 0):
            return False
    return True


def inspect_docx(path, spec):
    with zipfile.ZipFile(path) as package:
        document = ET.fromstring(package.read('word/document.xml'))
        styles_xml = ET.fromstring(package.read('word/styles.xml'))
        numbering = ET.fromstring(package.read('word/numbering.xml')) if 'word/numbering.xml' in package.namelist() else None
        numbering_policy = _section_numbering_policy(package, document) if spec.get('section_page_number_policy') else None
    styles = {item.get(W + 'styleId'): item for item in styles_xml.findall('w:style', NS)}
    paragraphs = document.findall('.//w:p', NS)
    checks = {}
    if numbering_policy is not None:
        checks['section_page_number_policy'] = numbering_policy
    bodies = [p for p in paragraphs if _text(p) == spec['body']]
    body_ok = len(bodies) == 1
    if bodies:
        body = bodies[0]
        fonts = _run_property(body, styles, 'rFonts', 'eastAsia')
        sizes = _run_property(body, styles, 'sz')
        indent_points = _effective(body, styles, 'w:pPr/w:ind', 'firstLine')
        indent_chars = _effective(body, styles, 'w:pPr/w:ind', 'firstLineChars')
        body_ok = body_ok and bool(fonts) and all(value and ('仿宋' in value or 'fangsong' in value.lower()) for value in fonts)
        body_ok = body_ok and bool(sizes) and all(value == '24' for value in sizes)
        body_ok = body_ok and (indent_chars == '200' if indent_chars is not None else indent_points == '480')
    checks['body_fangsong_12pt_two_character_indent'] = bool(body_ok)
    headings = []
    sizes = [32, 30, 30, 28, 28, 24]
    for index, label in enumerate(spec['headings']):
        # TOC result paragraphs also contain heading text; retain actual outline
        # paragraphs so cached TOC content cannot impersonate a native heading.
        matches = [p for p in paragraphs if _text(p) == label and _effective(p, styles, 'w:pPr/w:outlineLvl') == str(index)]
        headings.append(matches[0] if len(matches) == 1 else None)
    checks['heading_sizes_and_outline_levels'] = all(p is not None and bool(_run_property(p, styles, 'sz')) and all(v == str(sizes[i]) for v in _run_property(p, styles, 'sz')) for i, p in enumerate(headings))
    checks['native_heading_numbering'] = not spec['numbered'] or all(p is not None and _heading_numbered(p, styles, numbering, i) for i, p in enumerate(headings[:spec.get('numbered_levels', len(headings))]))
    if spec['title']:
        checks['single_cover_title'] = _text(document).count(spec['title']) == 1
    fields = [x.text or '' for x in document.findall('.//w:instrText', NS)]
    fields += [x.get(W + 'instr', '') for x in document.findall('.//w:fldSimple', NS)]
    checks['native_toc_field'] = not spec['toc'] or any(re.search(r'\bTOC\b', field) for field in fields)
    if spec['table_rows']:
        tables = document.findall('.//w:tbl', NS)
        matches = [table for table in tables if [[_text(cell).strip() for cell in row.findall('w:tc', NS)] for row in table.findall('w:tr', NS)] == spec['table_rows']]
        checks['table_content_exact'] = len(matches) == 1
        checks['table_header_repeats'] = bool(matches) and matches[0].find('w:tr/w:trPr/w:tblHeader', NS) is not None
    combined = _text(document)
    checks['end_markers_preserved_once'] = all(combined.count(marker) == 1 for marker in spec['markers'])
    sections = []
    for section in document.findall('.//w:sectPr', NS):
        page_size = section.find('w:pgSz', NS)
        sections.append({
            'width_twips': page_size.get(W + 'w') if page_size is not None else None,
            'height_twips': page_size.get(W + 'h') if page_size is not None else None,
            'orientation': page_size.get(W + 'orient', 'portrait') if page_size is not None else None,
        })
    def a4_portrait(section):
        try:
            return (section['orientation'] == 'portrait'
                    and abs(float(section['width_twips']) - 11906) <= 4
                    and abs(float(section['height_twips']) - 16838) <= 4)
        except (TypeError, ValueError):
            return False
    checks['all_sections_a4_portrait'] = bool(sections) and all(a4_portrait(section) for section in sections)
    heading_details = [{
        'label': spec['headings'][i],
        'style_id': _value(paragraph, 'w:pPr/w:pStyle') if paragraph is not None else None,
        'effective_half_point_sizes': _run_property(paragraph, styles, 'sz') if paragraph is not None else [],
        'expected_half_point_size': sizes[i],
    } for i, paragraph in enumerate(headings)]
    return {'checks': checks, 'sections': sections, 'headings': heading_details,
            'bytes': Path(path).stat().st_size, 'sha256': sha256(path)}


def inspect_pdf(path, spec):
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        pages = [page.extract_text() or '' for page in pdf.pages]
        footers = [''.join(word['text'] for word in page.extract_words() if word['top'] > page.height - 80) for page in pdf.pages] if spec.get('section_page_number_policy') else None
        page_sizes = [{'width_points': float(page.width), 'height_points': float(page.height)} for page in pdf.pages]
    compact = re.sub(r'\s+', '', ''.join(pages))
    required = [spec['body'], *spec['headings'], *spec['markers']]
    if spec['title']:
        required.append(spec['title'])
    required.extend(row[1] for row in spec['table_rows'][1:])
    checks = {
        'all_pdf_pages_a4_portrait': bool(page_sizes) and all(abs(size['width_points'] - 595.28) <= 0.25 and abs(size['height_points'] - 841.89) <= 0.25 for size in page_sizes),
        'nonempty_native_pdf_pages': bool(pages) and all(text.strip() for text in pages),
        'pdf_required_content': all(re.sub(r'\s+', '', text) in compact for text in required),
        'pdf_end_marker_once': all(compact.count(marker) == 1 for marker in spec['markers']),
    }
    if footers is not None:
        checks['pdf_section_page_numbers'] = len(footers) >= 3 and footers == ['', 'i'] + [str(i) for i in range(1, len(footers)-1)]
    return {'checks': checks, 'footer_text': footers, 'pages': len(pages), 'page_sizes': page_sizes, 'bytes': Path(path).stat().st_size, 'sha256': sha256(path)}


def _candidate(deadline):
    def remaining():
        value = deadline - time.monotonic()
        if value <= 0:
            raise TimeoutError('Acceptance source fingerprint deadline expired')
        return value

    def git(*args):
        result = subprocess.run(['git', '-C', str(ROOT), *args], capture_output=True, text=True, timeout=min(10, remaining()))
        return result.stdout.strip() if result.returncode == 0 else None
    addin = ROOT / 'macos' / 'wps-jsapi-probe' / 'addin'
    files = sorted((ROOT / 'skills' / 'WPSComposer').rglob('*.py')) + sorted(addin.glob('*.js')) + [Path(__file__).resolve()]
    if (addin / 'asset-manifest.json').is_file():
        files.append(addin / 'asset-manifest.json')
    return {
        'commit': git('rev-parse', 'HEAD'), 'branch': git('branch', '--show-current'),
        'working_tree_status': git('status', '--porcelain'),
        'source_files': {str(path.relative_to(ROOT)).replace('\\', '/'): sha256(path) for path in files if remaining()},
    }


def run_acceptance(output_root, *, fixture='representative', engine='msoffice', timeout=600):
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('Acceptance timeout must be positive and finite')
    if engine not in {'msoffice', 'wps'}:
        raise ValueError('Acceptance engine must be explicitly msoffice or wps')
    content, spec = fixture_content(fixture)
    root = Path(output_root).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=False)
    deadline = time.monotonic() + timeout
    report = {'schema': 1, 'status': 'RUNNING', 'engine': engine, 'fixture': fixture,
              'platform': platform.platform(), 'python': sys.version,
              'checks': {}, 'artifacts': {}, 'errors': [],
              'native_ui_acceptance': 'NOT_RUN', 'timeout_seconds': timeout}

    def persist():
        temporary = root / 'report.tmp'
        temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        temporary.replace(root / 'report.json')

    def remaining():
        value = deadline - time.monotonic()
        if value <= 0:
            raise TimeoutError('Acceptance total deadline expired')
        return value

    persist()
    started = time.monotonic()
    try:
        source = root / 'source.md'
        source.write_text(content, encoding='utf-8')
        report['source'] = {'path': 'source.md', 'sha256': sha256(source)}
        report['candidate'] = _candidate(deadline)
        persist()
        from skills.WPSComposer import generate, convert_to_pdf
        docx = root / 'generated.docx'
        pdf = root / 'converted.pdf'
        result = generate(str(source), format='docx', output=str(docx), engine=engine, timeout=remaining(), open_result=False)
        report['checks']['public_generate_returned_requested_path'] = Path(result).resolve() == docx
        before = sha256(docx)
        report['artifacts']['docx'] = inspect_docx(docx, spec)
        persist()
        converted = convert_to_pdf(str(docx), str(pdf), engine=engine, timeout=remaining(), open_result=False)
        report['checks']['public_conversion_returned_requested_path'] = Path(converted).resolve() == pdf
        report['checks']['conversion_preserved_source_docx_bytes'] = sha256(docx) == before
        report['artifacts']['pdf'] = inspect_pdf(pdf, spec)
        persist()
        direct_pdf = root / 'direct-generated.pdf'
        direct_result = generate(str(source), format='pdf', output=str(direct_pdf), engine=engine, timeout=remaining(), open_result=False)
        report['checks']['public_direct_pdf_returned_requested_path'] = Path(direct_result).resolve() == direct_pdf
        report['artifacts']['direct_pdf'] = inspect_pdf(direct_pdf, spec)
        persist()
        refused_generate = refused_convert = False
        pdf_before = sha256(pdf)
        try:
            generate(str(source), format='docx', output=str(docx), engine=engine, timeout=remaining(), open_result=False)
        except FileExistsError:
            refused_generate = True
        try:
            convert_to_pdf(str(docx), str(pdf), engine=engine, timeout=remaining(), open_result=False)
        except FileExistsError:
            refused_convert = True
        report['checks']['generation_overwrite_refused_without_change'] = refused_generate and sha256(docx) == before
        report['checks']['conversion_overwrite_refused_without_change'] = refused_convert and sha256(pdf) == pdf_before
        report['checks']['markdown_source_unchanged'] = sha256(source) == report['source']['sha256']
        remaining()
        all_checks = list(report['checks'].values()) + [value for artifact in report['artifacts'].values() for value in artifact['checks'].values()]
        report['status'] = 'PASS' if all_checks and all(value is True for value in all_checks) else 'FAIL'
    except BaseException as exc:
        report['status'] = 'FAIL'
        report['errors'].append({'type': type(exc).__name__, 'message': str(exc), 'traceback': traceback.format_exc()})
    finally:
        report['elapsed_seconds'] = round(time.monotonic() - started, 3)
        persist()
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-root', type=Path, required=True)
    parser.add_argument('--fixture', choices=['smoke', 'representative'], default='representative')
    parser.add_argument('--engine', choices=['msoffice', 'wps'], default='msoffice')
    parser.add_argument('--timeout', type=float, default=600)
    args = parser.parse_args()
    report = run_acceptance(args.output_root, fixture=args.fixture, engine=args.engine, timeout=args.timeout)
    print(json.dumps({'status': report['status'], 'report': str(args.output_root.resolve() / 'report.json'), 'native_ui_acceptance': report['native_ui_acceptance']}, ensure_ascii=False))
    return 0 if report['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
