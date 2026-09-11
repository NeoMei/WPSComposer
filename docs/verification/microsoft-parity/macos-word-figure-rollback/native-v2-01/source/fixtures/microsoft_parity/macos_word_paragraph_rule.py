"""Source-bound public paragraph-rule acceptance. No native work without --execute."""
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
from fixtures.microsoft_parity.macos_word_recovery import inventory, close_sentinel
from fixtures.microsoft_parity.macos_word_fields import retain_sources

PREFIX = 'PARAGRAPH RULE 中文😀 — unrelated existing body'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
SOURCES = [Path(__file__),
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_rules.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_script.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_runtime.py',
           ROOT / 'fixtures/microsoft_parity/macos_word_recovery.py',
           ROOT / 'fixtures/microsoft_parity/macos_word_fields.py',
           ROOT / 'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py',
           ROOT / 'skills/WPSComposer/scripts/writer.py']


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def readback():
    return [
        'set ruleRange to text object of paragraph 1 of boundDoc',
        'set followingRange to text object of paragraph 2 of boundDoc',
        'set ownBorder to get border ruleRange which border border bottom',
        'set followingBorder to get border followingRange which border border bottom',
        'set fmt to paragraph format of followingRange',
        'set followingStyleName to name local of style of followingRange as text',
        'set expectedBodyStyleName to name local of Word style (style body text) of boundDoc as text',
        "set followingIsBodyText to ((current application's NSString's stringWithString:followingStyleName)'s isEqualToString:expectedBodyStyleName) as boolean",
        'set nativeRows to {{"public-paragraph-rule",'
        'content of ruleRange as text,content of followingRange as text,'
        '(alignment of paragraph format of ruleRange is align paragraph center),'
        '(line style of ownBorder is line style single),'
        '(line width of ownBorder is line width75 point),'
        '(color of ownBorder is {49344,49344,49344}),'
        '(line style of followingBorder is line style none),'
        'followingIsBodyText,'
        '(alignment of fmt is align paragraph center),'
        'space before of fmt,space after of fmt,first line indent of fmt,line spacing of fmt}}',
    ]


def native_valid(rows):
    return (isinstance(rows, list) and len(rows) == 1 and isinstance(rows[0], list)
            and len(rows[0]) == 14
            and rows[0][:3] == ['public-paragraph-rule', PREFIX+' \r', 'FOLLOWING\r']
            and all(value is True for value in rows[0][3:10])
            and all(type(value) in (int, float) for value in rows[0][10:])
            and rows[0][10:] == [6, 9, 12, 18])


def xml_checks(xml, styles_xml):
    paragraphs = ET.fromstring(xml).findall('./'+W+'body/'+W+'p')
    if len(paragraphs) != 2:
        return {'native_border_xml': False, 'prefix_text_preserved': False, 'following_style_spacing_xml': False}
    current, following = paragraphs
    styles = ET.fromstring(styles_xml).findall(W+'style')
    def prop(paragraph, name, key):
        element = paragraph.find('./'+W+'pPr/'+W+name)
        return element.get(W+key) if element is not None else None
    following_style_id = prop(following, 'pStyle', 'val')
    following_styles = [style for style in styles if following_style_id is not None and style.get(W+'styleId') == following_style_id]
    following_is_body = (len(following_styles) == 1
        and following_styles[0].get(W+'type') == 'paragraph'
        and len(following_styles[0].findall(W+'name')) == 1
        and following_styles[0].find(W+'name').get(W+'val') == 'Body Text')
    border = current.find('./'+W+'pPr/'+W+'pBdr/'+W+'bottom')
    after = following.find('./'+W+'pPr/'+W+'pBdr/'+W+'bottom')
    text = lambda paragraph: ''.join(n.text or '' for n in paragraph.iter(W+'t'))
    return {'native_border_xml': (border is not None and border.get(W+'val') == 'single'
                and border.get(W+'sz') == '6' and border.get(W+'color','').upper() == 'C0C0C0'
                and prop(current,'jc','val') == 'center'
                and (after is None or after.get(W+'val') in ('nil','none'))),
            'prefix_text_preserved': text(current) == PREFIX+' ' and text(following) == 'FOLLOWING',
            'following_style_spacing_xml': (following_is_body
                and prop(following,'jc','val') == 'center'
                and prop(following,'spacing','before') == '120'
                and prop(following,'spacing','after') == '180' and prop(following,'spacing','line') == '360'
                and prop(following,'spacing','lineRule') == 'exact' and prop(following,'ind','firstLine') == '240')}


def pdf_rule_matches(drawing, page_number, page_rect):
    """Recognize this fixture's 0.75pt C0C0C0 rule in its expected A4 position."""
    def number(value):
        return type(value) in (int, float) and math.isfinite(value)

    def gray(value):
        return (isinstance(value, (list, tuple)) and len(value) == 3
                and all(number(c) and abs(c - 192/255) < 0.01 for c in value))

    if (page_number != 1 or abs(page_rect.width - 595.2) > 1
            or abs(page_rect.height - 841.92) > 1):
        return False
    items = drawing.get('items', [])
    if len(items) != 1:
        return False
    item = items[0]
    if drawing.get('type') == 'f' and len(item) == 3 and item[0] == 're':
        rect = item[1]
        if (not gray(drawing.get('fill')) or drawing.get('fill_opacity') != 1
                or not 0.5 <= rect.height <= 1.0):
            return False
        x0, x1, y = rect.x0, rect.x1, (rect.y0 + rect.y1)/2
    elif drawing.get('type') == 's' and len(item) == 3 and item[0] == 'l':
        start, end = item[1:]
        width = drawing.get('width')
        if (not gray(drawing.get('color')) or drawing.get('stroke_opacity') != 1
                or not number(width) or not 0.5 <= width <= 1.0
                or drawing.get('dashes') not in (None, '[] 0')
                or abs(start.y - end.y) > 0.05):
            return False
        x0, x1, y = min(start.x, end.x), max(start.x, end.x), (start.y + end.y)/2
    else:
        return False
    return (all(number(v) for v in (x0, x1, y)) and 414 <= x1 - x0 <= 422
            and abs(x0 - 88.56) <= 2 and abs(x1 - 506.64) <= 2
            and abs(y - 97.32) <= 2)


def artifacts(output):
    with ZipFile(output / 'paragraph-rule.docx') as package:
        xml = package.read('word/document.xml')
        styles_xml = package.read('word/styles.xml')
    (output / 'document.xml').write_bytes(xml)
    (output / 'styles.xml').write_bytes(styles_xml)
    import fitz
    with fitz.open(output / 'paragraph-rule.pdf') as pdf:
        text = '\n'.join(page.get_text() for page in pdf)
        rule_drawn = any(pdf_rule_matches(drawing, page_index + 1, page.rect)
                         for page_index, page in enumerate(pdf) for drawing in page.get_drawings())
        pdf[0].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(output/'page-1.png')
    (output/'pdf-text.txt').write_text(text)
    return dict(xml_checks(xml, styles_xml), pdf_text_visible='PARAGRAPH RULE' in text and 'FOLLOWING' in text,
                pdf_rule_drawn=rule_drawn)


def sentinel_preimage(rows, name, token):
    """Keep Word's actual full-name representation; verify identity/state/hash."""
    matches = [row for row in rows if isinstance(row, list) and row and row[0] == name]
    if len(matches) != 1 or not token:
        raise ValueError('Sentinel identity is not unique or token is empty')
    row = matches[0]
    if (len(row) != 4 or not isinstance(row[1], str) or row[2] is not False
            or row[3] != hashlib.sha256((token+'\r').encode()).hexdigest()):
        raise ValueError('Sentinel state or exact text hash is unverified')
    return list(row)


def close_after_owned(session, output, report, name, token):
    """Prove owned close and exact unsaved sentinel before independently closing it."""
    report['sentinel_cleanup_attempted'] = True
    if not session._closed or session._quarantined:
        raise RuntimeError('Owned document close is unverified; sentinel retained')
    observed = inventory(output, 'after-owned-close')
    report['inventory_after_owned_close'] = observed
    if ([r for r in observed if r[0] == name] != [report['sentinel_before']]
            or [r for r in observed if r[0] != name] != report['inventory_before']):
        raise RuntimeError('Post-owned-close inventory changed; sentinel retained')
    report['checks']['owned_closed_with_sentinel_preserved'] = True
    report['sentinel_close'] = close_sentinel(output, name, token)
    report['sentinel_closed'] = True
    report['inventory_after_sentinel_close'] = inventory(output, 'after-sentinel-close')
    report['checks']['inventory_after_sentinel_close_preserved'] = report['inventory_after_sentinel_close'] == report['inventory_before']
    if not report['checks']['inventory_after_sentinel_close_preserved']:
        raise RuntimeError('Final inventory changed')


def run(output):
    report = {'status': 'FAIL', 'checks': {}, 'scope': 'public method native/artifact acceptance; UI edit/undo remains separate',
              'source_hashes': retain_sources(output, SOURCES)}
    session = None; reopened = None; sentinel_name = None; token = None
    try:
        if not hasattr(MacWordSession, 'add_paragraph_horizontal_line'):
            raise RuntimeError('Public paragraph-rule forwarding is not integrated; native execution refused')
        report['dictionary_sha256'] = sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before'] = inventory(output, 'inventory-before')
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence = True
            try:
                token = 'PARAGRAPH RULE SENTINEL 中文😀 ' + uuid4().hex
                sentinel_name = session._execute(['set sentinelDoc to make new document',
                    f'set content of text object of sentinelDoc to {apple_string(token)}',
                    'set nativeRows to {{name of sentinelDoc as text}}'])[0][0]
                report['sentinel_name'] = sentinel_name
                created_inventory = inventory(output, 'before-owned-write')
                report['sentinel_before'] = sentinel_preimage(created_inventory, sentinel_name, token)
                report['word_version'] = session._execute(['set nativeRows to {{version as text}}'])[0][0]
                session._execute_structural([
                    f'set content of text object of boundDoc to {apple_string(PREFIX)}',
                    'set beforeRange to text object of paragraph 1 of boundDoc',
                    'set style of beforeRange to Word style (style body text) of boundDoc',
                    'set space before of paragraph format of beforeRange to 6',
                    'set space after of paragraph format of beforeRange to 9',
                    'set first line indent of paragraph format of beforeRange to 12',
                    'set line spacing rule of paragraph format of beforeRange to line space exactly',
                    'set line spacing of paragraph format of beforeRange to 18',
                    'set nativeRows to {{"setup",true}}'])
                session.save_docx(output/'preimage.docx')
                report['preimage_sha256'] = sha(output/'preimage.docx')
                report['public_result'] = session.add_paragraph_horizontal_line()
                session._execute_structural(session._position('end') + [
                    'set followingInsertion to create range boundDoc start insertionPoint end insertionPoint',
                    'set content of followingInsertion to "FOLLOWING"', 'set nativeRows to {{"following",true}}'])
                report['native_rows'] = session._execute(readback())
                report['checks']['native_readback'] = native_valid(report['native_rows'])
                if not report['checks']['native_readback']:
                    raise AssertionError('Public paragraph rule or inherited following format not verified')
                session.save_docx(output/'paragraph-rule.docx'); session.export_pdf(output/'paragraph-rule.pdf')
            except BaseException:
                report['primary_failure'] = traceback.format_exc()
                raise
        # Owned context has exited; retain evidence including its close script/log.
        shutil.copytree(session.staging_root, output/'native-runtime', dirs_exist_ok=True)
        try:
            close_after_owned(session, output, report, sentinel_name, token)
            sentinel_name = None
        except BaseException:
            report['sentinel_cleanup_failure'] = traceback.format_exc()
            raise
        digest = sha(output/'paragraph-rule.docx')
        with MacWordSession.open_document(output/'paragraph-rule.docx', read_only=True, visible=False) as reopened:
            reopened._retain_evidence = True
            report['reopened_rows'] = reopened._execute(readback())
            report['checks']['reopen_native_readback'] = native_valid(report['reopened_rows'])
        shutil.copytree(reopened.staging_root, output/'reopen-runtime', dirs_exist_ok=True)
        report['inventory_after_reopen_close'] = inventory(output, 'after-reopen-close')
        report['checks']['reopen_owned_closed_inventory_preserved'] = reopened._closed and report['inventory_after_reopen_close'] == report['inventory_before']
        report['checks']['source_unchanged_after_reopen'] = digest == sha(output/'paragraph-rule.docx')
        report['checks']['preimage_file_unchanged'] = report['preimage_sha256'] == sha(output/'preimage.docx')
        report['checks'].update(artifacts(output))
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix in ('.docx','.pdf','.png','.xml')}
        report['status'] = 'PASS' if all(report['checks'].values()) and not sentinel_name else 'FAIL'
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        # Failure cleanup is independent and only after verified owned close.
        # Never overwrite the primary failure or issue events after quarantine.
        if session and sentinel_name and not report.get('sentinel_cleanup_attempted'):
            try:
                close_after_owned(session, output, report, sentinel_name, token)
                sentinel_name = None
            except BaseException:
                report['sentinel_cleanup_failure'] = traceback.format_exc()
        if report.get('sentinel_closed'):
            sentinel_name = None
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root, output/label, dirs_exist_ok=True)
        if session:
            report['quarantined'] = session._quarantined or bool(reopened and reopened._quarantined)
            report['recovery_targets'] = {'sentinel_name':sentinel_name, 'owned_document':session._bound_path,
                                          'staging_root':str(session.staging_root)}
        if sentinel_name or report.get('sentinel_cleanup_failure'):
            report['status'] = 'FAIL'
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Refusing native Word mutation without --execute', file=sys.stderr)
        return 2
    output = args.output.resolve(); output.mkdir(parents=True, exist_ok=False)
    return 0 if run(output)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
