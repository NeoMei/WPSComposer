"""Opt-in Flat OPC *input* import hypothesis; never a production document writer.

Word.sdef has insert file(at:text range), format flat document (19), and
open format xmldocument serialized (14). Import format auto-detection remains
unproved. Only Word imports, materializes styles, edits and saves output DOCX.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_feasibility import (
    CLONE, W, SOURCES as OLD_SOURCES, MacWordSession, apple_string, sha,
    inventory, retain_sources, sentinel_preimage, close_after_owned,
    seed_commands, canonical, paragraph_style_by_name, clone_xml_valid,
    package_styles, style_definition, diagnostic_readback, scalar_copy_verdict,
    clone_readback, native_valid, diagnostic_rows_valid,
)

PKG = '{http://schemas.microsoft.com/office/2006/xmlPackage}'
REL = '{http://schemas.openxmlformats.org/package/2006/relationships}'
REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
STYLE_ID = 'WPSCImportedDetachedHeading'
IMPORTED = 'IMPORTED APPEARANCE 中文 العربية 😀'
SOURCE = 'SOURCE APPEARANCE 中文 العربية 😀'
FRESH = 'FRESH STYLE 中文 العربية 😀'
SEED = SOURCE + '\rREPLACE\rSUFFIX\r'
SOURCES = [Path(__file__), ROOT/'tests/msoffice/test_macos_word_heading_import.py', *OLD_SOURCES]


def flat_opc_fragment(native_styles):
    """Carry exact native-saved rPr/pPr into a distinct input style, not output."""
    source_xml = ET.fromstring(native_styles)
    styles = source_xml.findall(W+'style')
    source = paragraph_style_by_name(styles, 'heading 1')
    normal = paragraph_style_by_name(styles, 'Normal')
    if (source is None or normal is None or source.find(W+'rPr') is None
            or source.find(W+'pPr') is None
            or source.find(W+'rPr').find(W+'szCs') is None
            or any(s.get(W+'styleId') == STYLE_ID or
                   (s.find(W+'name') is not None and s.find(W+'name').get(W+'val') == CLONE)
                   for s in styles)):
        raise ValueError('Seed style identity, complete appearance or discriminating szCs is absent')
    package = ET.Element(PKG+'package')
    def part(name, content_type, element):
        node = ET.SubElement(package, PKG+'part', {PKG+'name': name, PKG+'contentType': content_type})
        ET.SubElement(node, PKG+'xmlData').append(element)
    def relationships(kind, target):
        node = ET.Element(REL+'Relationships')
        ET.SubElement(node, REL+'Relationship', {'Id': 'rId1', 'Type': REL_TYPE+kind, 'Target': target})
        return node
    part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml', relationships('officeDocument', 'word/document.xml'))
    imported_styles = ET.Element(W+'styles')
    defaults = source_xml.find(W+'docDefaults')
    if defaults is not None:
        imported_styles.append(deepcopy(defaults))
    imported_styles.append(deepcopy(normal))
    clone = ET.SubElement(imported_styles, W+'style', {W+'type': 'paragraph', W+'customStyle': '1', W+'styleId': STYLE_ID})
    ET.SubElement(clone, W+'name', {W+'val': CLONE})
    ET.SubElement(clone, W+'basedOn', {W+'val': normal.get(W+'styleId')})
    ET.SubElement(clone, W+'next', {W+'val': normal.get(W+'styleId')})
    clone.append(deepcopy(source.find(W+'pPr')))
    clone.append(deepcopy(source.find(W+'rPr')))
    part('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', imported_styles)
    document = ET.Element(W+'document')
    paragraph = ET.SubElement(ET.SubElement(document, W+'body'), W+'p')
    ET.SubElement(ET.SubElement(paragraph, W+'pPr'), W+'pStyle', {W+'val': STYLE_ID})
    ET.SubElement(ET.SubElement(paragraph, W+'r'), W+'t').text = IMPORTED
    part('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', document)
    part('/word/_rels/document.xml.rels', 'application/vnd.openxmlformats-package.relationships+xml', relationships('styles', 'styles.xml'))
    return ET.tostring(package, encoding='utf-8', xml_declaration=True)


def existing_styles_preserved(before, after):
    def index(xml):
        result = {}
        for style in ET.fromstring(xml).findall(W+'style'):
            key = style.get(W+'styleId')
            if not key or key in result:
                return None
            result[key] = (tuple(sorted(style.attrib.items())), tuple(canonical(c) for c in style if c.tag != W+'rsid'))
        return result
    old, new = index(before), index(after)
    return old is not None and new is not None and all(new.get(k) == v for k, v in old.items())


def fresh_paragraph_xml_valid(document_xml, styles_xml):
    """Require a unique style-only paragraph, including unexposed szCs/pPr."""
    styles = ET.fromstring(styles_xml).findall(W+'style')
    clone = paragraph_style_by_name(styles, CLONE)
    if clone is None or clone.find(W+'link') is not None:
        return False
    paragraphs = [p for p in ET.fromstring(document_xml).iter(W+'p')
                  if ''.join(t.text or '' for t in p.iter(W+'t')) == FRESH]
    if len(paragraphs) != 1:
        return False
    paragraph = paragraphs[0]
    props = paragraph.findall(W+'pPr')
    if len(props) != 1 or len(props[0]) != 1:
        return False
    style = props[0][0]
    if style.tag != W+'pStyle' or style.get(W+'val') != clone.get(W+'styleId'):
        return False
    # No direct run/paragraph-mark appearance can hide a complex-script gap.
    return all(len(rpr) == 0 and not rpr.attrib for rpr in paragraph.iter(W+'rPr'))


def fresh_package_valid(path):
    with ZipFile(path) as package:
        return fresh_paragraph_xml_valid(package.read('word/document.xml'), package.read('word/styles.xml'))


def replacement_expectation(preimage, start, end, inserted):
    """Reference splice in native UTF-16 units; preserve every outside character."""
    if (not isinstance(preimage, str) or not isinstance(inserted, str)
            or type(start) is not int or type(end) is not int):
        raise ValueError('Invalid native replacement preimage')
    encoded = preimage.encode('utf-16-le')
    if not 0 <= start <= end <= len(encoded)//2:
        raise ValueError('Invalid native UTF-16 bounds')
    try:
        prefix = encoded[:start*2].decode('utf-16-le')
        target = encoded[start*2:end*2].decode('utf-16-le')
        suffix = encoded[end*2:].decode('utf-16-le')
    except UnicodeDecodeError:
        raise ValueError('Native bounds split a surrogate pair') from None
    if target != 'REPLACE\r':
        raise ValueError('Native replacement target changed')
    return prefix + inserted + suffix


def replacement_snapshot():
    return ['set importRange to text object of paragraph 2 of boundDoc',
            'set nativeRows to {{"preimage",content of text object of boundDoc as text,start of content of importRange,end of content of importRange}}']


def import_commands(path):
    return [
        'set importRange to text object of paragraph 2 of boundDoc',
        'set importStart to start of content of importRange',
        'set importEnd to end of content of importRange',
        'if content of importRange is not "REPLACE" & return then error "WPSC_IMPORT_PREIMAGE"',
        f'insert file at importRange file name {apple_string(str(path))} confirm conversions false link false',
        'set nativeRows to {{"import",importStart,importEnd,content of text object of boundDoc as text}}',
    ]


def seed():
    # Reuse the proven run-04 font seed, without creating/applying its clone.
    lines = seed_commands('font-properties')
    lines = [line for line in lines if not any(token in line for token in
             ('set detachedStyle', 'set base style of detachedStyle',
              'set content of text object of boundDoc', 'set style of text object of paragraph', 'set nativeRows'))]
    return lines + [f'set content of text object of boundDoc to {apple_string(SEED)}',
                    'set style of text object of paragraph 1 of boundDoc to sourceStyle',
                    'set nativeRows to {{"stage",true}}']


def run(output):
    report = {'status': 'FAIL', 'scope': 'Native file-import hypothesis only; public method remains absent',
              'checks': {}, 'source_hashes': retain_sources(output, SOURCES)}
    session = reopened = None
    sentinel_name = token = None
    def flush():
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    def check(name, value):
        report['checks'][name] = bool(value)
        flush()
        if not value:
            raise AssertionError('Native import check failed: '+name)
    def execute(owner, label, lines):
        report['current_phase'] = label
        flush()
        rows = owner._execute(lines)
        report.setdefault('raw_rows', {})[label] = rows
        flush()
        return rows
    def observe(owner, label):
        rows = execute(owner, label, diagnostic_readback())
        for key, value in scalar_copy_verdict(font_preimage, rows).items():
            check(label+'_'+key, value)
        check(label+'_full_font_paragraph', native_valid('detached-properties', execute(owner, label+'_paragraph', clone_readback('detached-properties'))))
    try:
        report['dictionary_sha256'] = sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before'] = inventory(output, 'before')
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence = True
            token = 'IMPORT SENTINEL '+uuid4().hex+' 中文😀'
            rows = execute(session, 'sentinel', ['set sentinelDoc to make new document', f'set content of text object of sentinelDoc to {apple_string(token)}', 'set nativeRows to {{name of sentinelDoc as text}}'])
            if len(rows) != 1 or len(rows[0]) != 1 or not isinstance(rows[0][0], str):
                raise ValueError('Invalid sentinel acknowledgement')
            sentinel_name = rows[0][0]
            report['sentinel_name'] = sentinel_name
            report['sentinel_before'] = sentinel_preimage(inventory(output, 'sentinel-before'), sentinel_name, token)
            report['word_version'] = execute(session, 'version', ['set nativeRows to {{version as text}}'])
            check('seed_ack', execute(session, 'seed', seed()) == [['stage', True]])
            session.save_docx(output/'source-preimage.docx')
            preimage_hash = sha(output/'source-preimage.docx')
            preimage = package_styles(output/'source-preimage.docx')
            (output/'source-styles.xml').write_bytes(preimage)
            fragment = flat_opc_fragment(preimage)
            (output/'fragment.xml').write_bytes(fragment)
            # Word's container-local file access is reused, with exact input bytes.
            native_fragment = session.staging_root/'heading-input.xml'
            native_fragment.write_bytes(fragment)
            # Native font preimage uses Heading1 for both columns; no clone yet.
            before_commands = [line.replace(f'Word style {apple_string(CLONE)}', 'Word style (style heading1)') for line in diagnostic_readback()]
            font_preimage = execute(session, 'source_font_preimage', before_commands)
            check('source_font_preimage_valid', diagnostic_rows_valid(font_preimage) and all(r[3] for r in font_preimage))
            before = execute(session, 'replacement_preimage', replacement_snapshot())
            check('replacement_preimage_shape', len(before) == 1 and len(before[0]) == 4 and before[0][0] == 'preimage')
            _, native_text, start, end = before[0]
            expected = replacement_expectation(native_text, start, end, IMPORTED+'\r')
            rows = execute(session, 'native_import', import_commands(native_fragment))
            check('exact_range_replacement', rows == [['import', start, end, expected]])
            observe(session, 'after_import')
            session.save_docx(output/'after-import.docx')
            styles = package_styles(output/'after-import.docx')
            check('all_existing_styles_preserved', existing_styles_preserved(preimage, styles))
            check('complete_imported_style_xml', clone_xml_valid(styles, 'detached-properties'))
            # A new plain paragraph must pick up the imported style without copied runs.
            lines = [f'set detachedStyle to Word style {apple_string(CLONE)} of boundDoc',
                     'set p to (end of content of text object of boundDoc) - 1',
                     'set freshRange to create range boundDoc start p end p',
                     f'set content of freshRange to {apple_string(FRESH)} & return',
                     'set freshRange to text object of paragraph 4 of boundDoc',
                     'set style of freshRange to detachedStyle', 'reset font object of freshRange',
                     'reset paragraph format of freshRange', 'set nativeRows to {{"stage",true}}']
            check('fresh_ack', execute(session, 'fresh', lines) == [['stage', True]])
            fresh_commands = [line.replace('set sourceFont to font object of sourceStyle', 'set sourceFont to font object of text object of paragraph 4 of boundDoc') for line in diagnostic_readback()]
            rows = execute(session, 'fresh_font', fresh_commands)
            check('fresh_font_equal', all(scalar_copy_verdict(font_preimage, rows).values()))
            session.save_docx(output/'before-challenge.docx')
            check('fresh_style_only_xml', fresh_package_valid(output/'before-challenge.docx'))
            clone_before = style_definition(package_styles(output/'before-challenge.docx'), CLONE)
            source_size = next(r[1][1] for r in font_preimage if r[0] == 'font:font size')
            challenge = ['set sourceStyle to Word style (style heading1) of boundDoc', f'set font size of font object of sourceStyle to {source_size+11}', 'set nativeRows to {{"stage",true}}']
            check('challenge_ack', execute(session, 'source_challenge', challenge) == [['stage', True]])
            session.save_docx(output/'source-challenged.docx')
            check('fresh_style_only_after_challenge', fresh_package_valid(output/'source-challenged.docx'))
            check('clone_independent', style_definition(package_styles(output/'source-challenged.docx'), CLONE) == clone_before)
            rows = execute(session, 'fresh_after_challenge', fresh_commands)
            check('fresh_independent', all(scalar_copy_verdict(font_preimage, rows).values()))
            challenge[1] = f'set font size of font object of sourceStyle to {source_size}'
            check('restore_ack', execute(session, 'restore_source', challenge) == [['stage', True]])
            observe(session, 'after_restore')
            session.save_docx(output/'heading.docx')
            session.export_pdf(output/'heading.pdf')
            final_styles = package_styles(output/'heading.docx')
            (output/'styles.xml').write_bytes(final_styles)
            check('source_fully_restored', existing_styles_preserved(preimage, final_styles))
            check('complete_final_style_xml', clone_xml_valid(final_styles, 'detached-properties'))
            check('fresh_style_only_final', fresh_package_valid(output/'heading.docx'))
        close_after_owned(session, output, report, sentinel_name, token)
        sentinel_name = None
        digest = sha(output/'heading.docx')
        with MacWordSession.open_document(output/'heading.docx', read_only=True, visible=False) as reopened:
            reopened._retain_evidence = True
            observe(reopened, 'reopen')
            rows = execute(reopened, 'reopen_fresh', fresh_commands)
            check('reopen_fresh_equal', all(scalar_copy_verdict(font_preimage, rows).values()))
        check('reopen_bytes_preserved', digest == sha(output/'heading.docx'))
        check('fresh_style_only_reopened_package', fresh_package_valid(output/'heading.docx'))
        check('preimage_bytes_preserved', preimage_hash == sha(output/'source-preimage.docx'))
        check('input_bytes_preserved', (output/'fragment.xml').read_bytes() == fragment and native_fragment.read_bytes() == fragment)
        import fitz
        with fitz.open(output/'heading.pdf') as pdf:
            text = '\n'.join(page.get_text() for page in pdf)
            for i, page in enumerate(pdf):
                page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(output/f'page-{i+1}.png')
        (output/'pdf-text.txt').write_text(text)
        check('pdf_text_present', all(marker in text for marker in ('SOURCE APPEARANCE', 'IMPORTED APPEARANCE', 'FRESH STYLE', 'SUFFIX')))
        report['inventory_final'] = inventory(output, 'final')
        check('inventory_preserved', report['inventory_final'] == report['inventory_before'])
        check('sources_unchanged', all(sha(ROOT/path) == value for path, value in report['source_hashes'].items()))
        report['status'] = 'PASS'
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and sentinel_name and not report.get('sentinel_cleanup_attempted'):
            try:
                close_after_owned(session, output, report, sentinel_name, token)
                sentinel_name = None
            except BaseException:
                report['cleanup_failure'] = traceback.format_exc()
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root, output/label, dirs_exist_ok=True)
        if sentinel_name or report.get('cleanup_failure'):
            report['status'] = 'FAIL'
        report['remaining_sentinel'] = sentinel_name
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix in ('.docx', '.pdf', '.xml', '.png')}
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Native import probe requires explicit --execute', file=sys.stderr)
        return 2
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    return 0 if run(output)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
