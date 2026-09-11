"""Isolated Word rules / WordArt feasibility probe. Native work requires --execute.

Each primitive has a separate owned private document and retained report. This
fixture does not enable production support. Syntax compilation is not native proof.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from fixtures.microsoft_parity.macos_word_references import inventory_without_launch

CASES = ('inline', 'paragraph', 'wordart')
ART_TEXT = 'WORDART 中文😀'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
V = '{urn:schemas-microsoft-com:vml}'
O = '{urn:schemas-microsoft-com:office:office}'
SOURCES = [Path(__file__), ROOT / 'skills/WPSComposer/scripts/writer.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_script.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_runtime.py',
           ROOT / 'fixtures/microsoft_parity/macos_word_references.py']


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def readback_commands(case):
    if case == 'inline':
        return ['set ownLine to inline shape 1 of boundDoc',
                'set nativeRows to {{"inline",count inline shapes of boundDoc,(inline shape type of ownLine is inline shape horizontal line),width of ownLine,height of ownLine}}']
    if case == 'paragraph':
        return ['set ruleRange to text object of paragraph 2 of boundDoc',
                'set followingRange to text object of paragraph 3 of boundDoc',
                'set ownBorder to get border ruleRange which border border bottom',
                'set followingBorder to get border followingRange which border border bottom',
                'set nativeRows to {{"paragraph",content of ruleRange as text,(alignment of paragraph format of ruleRange is align paragraph center),(line style of ownBorder is line style single),(line width of ownBorder is line width75 point),(color of ownBorder is {49344,49344,49344}),content of followingRange as text,(line style of followingBorder is line style none)}}']
    if case == 'wordart':
        return ['set ownArt to shape 1 of boundDoc',
                'set artFormat to word art format of ownArt',
                'set nativeRows to {{"wordart",count shapes of boundDoc,(shape type of ownArt is shape type word art),word art text of artFormat as text,font name of artFormat as text,font size of artFormat,bold of artFormat,italic of artFormat,left position of ownArt,top of ownArt,(preset word art effect of artFormat is wordart format1),width of ownArt,height of ownArt}}']
    raise ValueError('Unknown case')


def build_probe_commands(case):
    if case not in CASES:
        raise ValueError('Unknown case')
    lines = ['set p to (end of content of text object of boundDoc) - 1',
             'set insertionRange to create range boundDoc start p end p']
    if case == 'inline':
        lines += ['set beforeCount to count inline shapes of boundDoc',
                  'make new standard inline horizontal line at insertionRange',
                  'if (count inline shapes of boundDoc) is not beforeCount + 1 then error "WPSC_RULE_COUNT_DELTA_FAILED"']
    elif case == 'paragraph':
        lines += ['set content of insertionRange to " " & return & "FOLLOWING"',
                  'set ruleRange to create range boundDoc start p end (p + 2)',
                  'set alignment of paragraph format of ruleRange to align paragraph center',
                  'set ownBorder to get border ruleRange which border border bottom',
                  'set line style of ownBorder to line style single',
                  'set line width of ownBorder to line width75 point',
                  'set color of ownBorder to {49344,49344,49344}',
                  'set followingRange to create range boundDoc start (p + 2) end (p + 11)',
                  'set followingBorder to get border followingRange which border border bottom',
                  'set line style of followingBorder to line style none']
    else:
        lines += ['set beforeCount to count shapes of boundDoc',
                  'make new word art at boundDoc with properties {anchor:insertionRange,word art text:' + apple_string(ART_TEXT) + ',font name:"Arial",font size:36,bold:false,italic:false,left position:200,top:400,preset word art effect:wordart format1}',
                  'if (count shapes of boundDoc) is not beforeCount + 1 then error "WPSC_WORDART_COUNT_DELTA_FAILED"',
                  'set ownArt to shape (count shapes of boundDoc) of boundDoc',
                  'if (shape type of ownArt) is not shape type word art then error "WPSC_WORDART_NATIVE_TYPE_FAILED"']
    return lines + readback_commands(case)


def _number(value):
    return type(value) in (int, float) and math.isfinite(value)


def verify_native_rows(case, rows):
    if not isinstance(rows, list) or len(rows) != 1 or not isinstance(rows[0], list):
        return False
    row = rows[0]
    if not row or row[0] != case:
        return False
    if case == 'inline':
        return (len(row) == 5 and type(row[1]) is int and row[1] == 1 and row[2] is True
                and all(_number(v) and v > 0 for v in row[3:]))
    if case == 'paragraph':
        return len(row) == 8 and row[1] == ' \r' and all(v is True for v in row[2:6]) and row[6] == 'FOLLOWING\r' and row[7] is True
    if case == 'wordart':
        return (len(row) == 13 and type(row[1]) is int and row[1] == 1 and row[2] is True
                and row[3:5] == [ART_TEXT, 'Arial'] and _number(row[5]) and row[5] == 36
                and row[6] is False and row[7] is False
                and all(_number(v) for v in row[8:10]) and row[8:10] == [200, 400]
                and row[10] is True and all(_number(v) and v > 0 for v in row[11:]))
    return False


def verify_xml(case, payload):
    doc = ET.fromstring(payload)
    if case == 'inline':
        return len([s for s in doc.iter(V + 'rect') if s.get(O + 'hr') in ('t', 'true')]) == 1
    if case == 'wordart':
        # VML textpath is editable native WordArt. A text box or image alone
        # cannot pass. Newer representations remain unverified until inspected.
        arts = [s for s in doc.iter(V + 'shape') if s.find(V + 'textpath') is not None]
        if len(arts) != 1:
            return False
        art = arts[0]; textpath = art.find(V + 'textpath')
        style = textpath.get('style', '').replace(' ', '').replace('"', '').replace("'", '').lower()
        geometry = art.get('style', '').replace(' ', '').lower()
        return (textpath.get('string') == ART_TEXT and 'font-family:arial' in style
                and 'font-size:36pt' in style and 'margin-left:200pt' in geometry
                and 'margin-top:400pt' in geometry and art.find(V + 'textbox') is None)
    if case == 'paragraph':
        paragraphs = doc.findall('./' + W + 'body/' + W + 'p')
        matches = []
        for index, p in enumerate(paragraphs[:-1]):
            if ''.join(p.itertext()) != ' ':
                continue
            border = p.find('./' + W + 'pPr/' + W + 'pBdr/' + W + 'bottom')
            align = p.find('./' + W + 'pPr/' + W + 'jc')
            following = paragraphs[index + 1]
            after = following.find('./' + W + 'pPr/' + W + 'pBdr/' + W + 'bottom')
            matches.append(border is not None and border.get(W + 'val') == 'single'
                           and border.get(W + 'sz') == '6' and border.get(W + 'color', '').upper() == 'C0C0C0'
                           and align is not None and align.get(W + 'val') == 'center'
                           and ''.join(following.itertext()) == 'FOLLOWING'
                           and (after is None or after.get(W + 'val') in ('nil', 'none')))
        return matches == [True]
    return False


def verify_artifacts(case, output):
    with ZipFile(output / (case + '.docx')) as package:
        xml = package.read('word/document.xml')
    (output / 'document.xml').write_bytes(xml)
    import fitz
    with fitz.open(output / (case + '.pdf')) as pdf:
        text = '\n'.join(page.get_text() for page in pdf)
        drawings = sum(len(page.get_drawings()) for page in pdf)
        pdf[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(output / 'page-1.png')
    (output / 'pdf-text.txt').write_text(text)
    return {'persisted_native_xml': verify_xml(case, xml),
            'pdf_anchor_visible': 'RULES PROBE' in text,
            'pdf_object_visible': ('FOLLOWING' in text and drawings > 0) if case == 'paragraph' else drawings > 0,
            'pdf_wordart_text': ('WORDART' in text) if case == 'wordart' else True}


def sentinel_commands(name, token, close=False):
    commands = [f'set sentinelDoc to document {apple_string(name)}',
                f'if (content of text object of sentinelDoc as text) is not {apple_string(token)} & return then error "WPSC_SENTINEL_CHANGED"',
                'if saved of sentinelDoc then error "WPSC_SENTINEL_SAVED"']
    if close:
        commands.append('close sentinelDoc saving no')
    return commands + ['set nativeRows to {{"sentinel",true}}']


def run_case(case, output):
    output.mkdir(parents=True, exist_ok=False)
    report = {'case': case, 'status': 'FAIL', 'checks': {}, 'source_hashes': {str(p.relative_to(ROOT)): sha(p) for p in SOURCES},
              'dictionary_hash': sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef'),
              'scope': 'probe only; public method and UI acceptance pending'}
    session = None; sentinel_name = None
    try:
        before = inventory_without_launch(output, 'inventory-before')
        report['inventory_before'] = before
        with MacWordSession.new_document(visible=False) as session:
            try:
                token = 'RULES SENTINEL 中文😀 ' + uuid4().hex
                sentinel_name = session._execute(['set sentinelDoc to make new document',
                    f'set content of text object of sentinelDoc to {apple_string(token)}',
                    'set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
                report['sentinel_identity'] = {'name': sentinel_name, 'text_hash': hashlib.sha256((token + '\r').encode()).hexdigest()}
                report['word_version'] = session._execute(['set nativeRows to {{version as text}}'])[0][0]
                session._execute_structural(['set content of text object of boundDoc to "RULES PROBE" & return', 'set nativeRows to {{"ok"}}'])
                rows = session._execute_structural(build_probe_commands(case))
                report['native_rows'] = rows
                report['checks']['native_readback'] = verify_native_rows(case, rows)
                if not report['checks']['native_readback']:
                    raise AssertionError('Native semantic readback mismatch')
                session.save_docx(output / (case + '.docx'))
                session.export_pdf(output / (case + '.pdf'))
                report['checks']['sentinel_unchanged'] = session._execute(sentinel_commands(sentinel_name, token)) == [['sentinel', True]]
            finally:
                if sentinel_name and not session._quarantined:
                    session._execute(sentinel_commands(sentinel_name, token, close=True)); sentinel_name = None
                if session.staging_root.exists():
                    shutil.copytree(session.staging_root, output / 'native-runtime', dirs_exist_ok=True)
        report['checks']['owned_closed'] = session._closed
        original_hash = sha(output / (case + '.docx'))
        with MacWordSession.open_document(output / (case + '.docx'), read_only=True, visible=False) as reopened:
            rows = reopened._execute(readback_commands(case))
            report['reopen_rows'] = rows
            report['checks']['reopen_native_readback'] = verify_native_rows(case, rows)
            shutil.copytree(reopened.staging_root, output / 'reopen-runtime', dirs_exist_ok=True)
        report['checks']['source_preserved_after_reopen'] = original_hash == sha(output / (case + '.docx'))
        report['checks'].update(verify_artifacts(case, output))
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix in ('.docx', '.pdf', '.png', '.xml')}
        report['status'] = 'PASS' if all(report['checks'].values()) else 'PARTIAL'
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output / 'failure.txt').write_text(traceback.format_exc())
        if session:
            report['quarantined'] = session._quarantined
            report['recovery_targets'] = {'sentinel_name': sentinel_name, 'owned_document': session._bound_path, 'staging_root': str(session.staging_root)}
            if session.staging_root.exists():
                shutil.copytree(session.staging_root, output / 'failed-runtime', dirs_exist_ok=True)
    finally:
        if not session or not session._quarantined:
            try:
                after = inventory_without_launch(output, 'inventory-after')
                report['inventory_after'] = after
                report['checks']['unrelated_inventory_preserved'] = report.get('inventory_before') == after
                if not report['checks']['unrelated_inventory_preserved']:
                    report['status'] = 'FAIL'
            except BaseException as exc:
                report['cleanup_error'] = repr(exc); report['status'] = 'FAIL'
        (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--case', choices=CASES, required=False)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Refusing native Word mutation without --execute', file=sys.stderr)
        return 2
    args.output.resolve().mkdir(parents=True, exist_ok=False)
    results = []
    for case in ([args.case] if args.case else CASES):
        result = run_case(case, args.output.resolve() / case)
        results.append(result)
        if result.get('quarantined') or result.get('recovery_targets', {}).get('sentinel_name'):
            break
    (args.output / 'report.json').write_text(json.dumps(results, ensure_ascii=False, indent=2))
    return 0 if all(r['status'] == 'PASS' for r in results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
