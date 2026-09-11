"""Guarded direct Mac Word notice acceptance; source-bound and --execute only.

No package writes: OOXML access is read-only acceptance. Each synthetic sentinel
is closed independently only after owned create/reopen contexts have closed.
"""
from __future__ import annotations

import argparse
from dataclasses import asdict
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_recovery import inventory, close_sentinel

SOURCES = [ROOT / relative for relative in (
    'fixtures/microsoft_parity/macos_word_degradation.py',
    'fixtures/microsoft_parity/macos_word_recovery.py',
    'fixtures/microsoft_parity/macos_word_fields.py',
    'skills/WPSComposer/__init__.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_degradation.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_fields.py',
    'skills/WPSComposer/scripts/msoffice/macos_word_references.py',
    'skills/WPSComposer/scripts/msoffice/macos_runtime.py',
    'skills/WPSComposer/scripts/msoffice/macos_script.py',
    'skills/WPSComposer/scripts/msoffice/errors.py',
    'skills/WPSComposer/scripts/artifact_transport.py',
    'skills/WPSComposer/scripts/writer.py',
    'skills/WPSComposer/scripts/longform/privacy.py',
)]
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def _on(element):
    return element is not None and element.get(W+'val', '1') not in ('0', 'false', 'off')


def _text(element):
    return ''.join(t.text or '' for t in element.iter(W+'t'))


def inspect_ooxml(path, display, following, *, block, table):
    """Verify persisted notice properties, row split, and an unstyled normal tail."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))
    paragraphs = list(root.iter(W+'p'))
    matches = [p for p in paragraphs if display in _text(p)]
    assert len(matches) == 1, 'Saved notice is missing or duplicated'
    notice = matches[0]
    start = _text(notice).index(display)
    end = start + len(display)
    offset = 0
    covered = 0
    for run in notice.iter(W+'r'):
        text = _text(run)
        right = offset + len(text)
        if offset < end and right > start:
            props = run.find(W+'rPr')
            assert props is not None and _on(props.find(W+'i')), 'Notice italic missing'
            color = props.find(W+'color')
            shading = props.find(W+'shd')
            assert color is not None and color.get(W+'val', '').upper() == '9C0006', 'Notice red missing'
            assert shading is not None and shading.get(W+'fill', '').upper() == 'FCE8E6', 'Notice shading missing'
            covered += min(right, end)-max(offset, start)
        offset = right
    assert covered == len(display), 'Not all notice characters verified'
    if block:
        props = notice.find(W+'pPr')
        assert props is not None
        spacing = props.find(W+'spacing')
        assert spacing is not None and spacing.get(W+'before', '0') == '0'
        assert spacing.get(W+'after') == '60', 'Notice 3pt after spacing missing'
        assert _on(props.find(W+'keepLines')), 'Notice keep-together missing'
        outline = props.find(W+'outlineLvl')
        assert outline is None or outline.get(W+'val') == '9', 'Notice is an outline heading'
    tables = list(root.iter(W+'tbl'))
    assert len(tables) == int(table), 'Unexpected residual or missing native table'
    if table:
        rows = tables[0].findall(W+'tr')
        assert len(rows) == 1 and len(rows[0].findall(W+'tc')) == 1
        assert _on(rows[0].find(W+'trPr/'+W+'cantSplit')), 'Notice row can split'
    normal = [p for p in paragraphs if _text(p) == following]
    assert len(normal) == 1, 'Following normal paragraph missing or duplicated'
    for run in normal[0].iter(W+'r'):
        props = run.find(W+'rPr')
        if props is None:
            continue
        assert not _on(props.find(W+'i')), 'Following normal text inherited italic'
        color, shading = props.find(W+'color'), props.find(W+'shd')
        assert color is None or color.get(W+'val', '').upper() != '9C0006'
        assert shading is None or shading.get(W+'fill', '').upper() != 'FCE8E6'
    shading = normal[0].find(W+'pPr/'+W+'shd')
    assert shading is None or shading.get(W+'fill', '').upper() != 'FCE8E6'
    return {'table_count': len(tables), 'notice_characters': covered,
            'normal_tail': following, 'block_paragraph_verified': block}


class Journal:
    def __init__(self, output):
        self.output = output
        self.data = {'passed': False, 'source_hashes': retain_sources(output, SOURCES),
                     'checks': {}, 'cases': {}}
        self.flush()

    def flush(self):
        path = self.output/'report.pending'
        path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding='utf-8')
        path.replace(self.output/'report.json')

    def check(self, case, name):
        self.data['checks'][case+'/'+name] = True
        self.flush()


def _body(session):
    return session._execute([
        'set nativeRows to {{content of text object of boundDoc as text,'
        'end of content of text object of boundDoc,count paragraphs of boundDoc,count tables of boundDoc}}',
    ])


def _format(session, start, end):
    return session._execute([
        f'set probeRange to create range boundDoc start {start} end {end}',
        'set nativeRows to {{start of content of probeRange,end of content of probeRange,'
        'content of probeRange as text,italic of font object of probeRange,'
        '(color of font object of probeRange is {40092, 0, 1542}),'
        '(background pattern color of shading of probeRange is {64764, 59624, 59110})}}',
    ])


def _retain_runtime(session, target):
    if session is not None and session.staging_root and session.staging_root.exists():
        shutil.copytree(session.staging_root, target, dirs_exist_ok=True)


def _case(journal, mode, nonempty):
    label = mode + ('-nonempty' if nonempty else '-empty')
    output = journal.output/label
    output.mkdir(exist_ok=False)
    case = journal.data['cases'][label] = {'passed': False}
    journal.flush()
    sentinel = None
    session = None
    try:
        starting = inventory(output, 'starting-inventory')
        prefix = 'PREFIX 中文😀' if nonempty else ''
        code = 'NOTICE_' + label.upper().replace('-', '_')
        display = '[' + code + (': 中文😀]' if mode == 'inline' else '] 中文😀')
        following = 'FOLLOWING-NORMAL-' + label
        with create_document('writer', engine='msoffice', visible=False) as session:
            try:
                sentinel_token = 'Notice sentinel 中文😀 ' + uuid4().hex
                sentinel = session._execute([
                    'set d to make new document',
                    f'set content of text object of d to {apple_string(sentinel_token)}',
                    'set nativeRows to {{name of d as text}}',
                ])[0][0]
                sentinel_before = next(row for row in inventory(output, 'before-owned-write') if row[0] == sentinel)
                case['sentinel_before'] = sentinel_before
                case['word_version'] = session._execute(['set nativeRows to {{version as text}}'])
                if prefix:
                    session._execute_structural([
                        f'set content of text object of boundDoc to {apple_string(prefix)}',
                        'set nativeRows to {{"prefix-ack"}}',
                    ])
                before = case['before'] = _body(session)
                assert before[0][0] == prefix+'\r' and before[0][3] == 0
                execute = session._execute
                recovery_module = sys.modules['skills.WPSComposer.scripts.msoffice.macos_word_recovery']
                module_rollback = recovery_module.rollback
                events = []
                # The helper calls the real recovery module. Count only after
                # its native preimage/deletion/postcondition ACK has returned.
                def observed_rollback(bound, token):
                    assert bound is session
                    result = module_rollback(bound, token)
                    assert execute([
                        'set nativeRows to {{content of text object of boundDoc as text,'
                        'end of content of text object of boundDoc,count paragraphs of boundDoc,count tables of boundDoc}}'
                    ]) == before
                    events.append('rollback-ack')
                    return result

                def controlled(lines, **kwargs):
                    lines = list(lines)
                    for i, line in enumerate(lines):
                        if line.startswith('set noticeTable to make new table at boundDoc '):
                            events.append('partial-table-submitted')
                            lines[i+1:i+1] = [
                                'set probeCell to get cell from table noticeTable row 1 column 1',
                                'set content of text object of probeCell to "PARTIAL-NOTICE-REMNANT"',
                                'error "WPSC_CONTROLLED_NOTICE_TABLE_FAILURE" number -2700',
                            ]
                            break
                    if any(line.startswith('set content of noticeRange to ') for line in lines):
                        assert events == ['partial-table-submitted', 'table-failed-ack', 'rollback-ack']
                        events.append('fallback-submitted')
                    result = execute(lines, **kwargs)
                    if isinstance(result, list) and result and result[0][0] == 'degradation-table-failed':
                        events.append('table-failed-ack')
                    return result

                try:
                    if mode == 'table-failure':
                        session._execute = controlled
                        recovery_module.rollback = observed_rollback
                    if mode == 'inline':
                        handle = session.add_inline_degradation(code, object(), '中文😀')
                        own_range = handle
                    else:
                        handle = session.add_degradation_notice(code, object(), '中文😀')
                        own_range = handle.Range
                finally:
                    session._execute = execute
                    recovery_module.rollback = module_rollback
                case['handle'] = asdict(handle)
                after = case['after_notice'] = _body(session)
                assert after[0][0].startswith(prefix) and after[0][0].endswith('\r')
                assert after[0][0].count(display) == 1 and 'PARTIAL-NOTICE-REMNANT' not in after[0][0]
                assert after[0][3] == int(mode == 'block')
                if mode == 'inline':
                    assert after[0][0] == prefix+display+'\r'
                    assert after[0][2] == before[0][2]
                if mode == 'table-failure':
                    assert events == ['partial-table-submitted', 'table-failed-ack', 'rollback-ack', 'fallback-submitted']
                    assert handle.table_index is None
                    case['recovery_events'] = events
                    journal.check(label, 'actual_partial_table_rollback_before_fallback_once')
                extent = len(display.encode('utf-16-le'))//2
                expected_format = [[own_range.Start, own_range.Start+extent, display, True, True, True]]
                case['native_notice_format'] = _format(session, own_range.Start, own_range.Start+extent)
                assert case['native_notice_format'] == expected_format
                journal.check(label, 'native_notice_utf16_format_and_semantic_handle')
                session.add_paragraph(following)
                body = case['saved_body'] = _body(session)
                # Table CR/BEL text expands native terminators. Resolve the
                # paragraph's own native start; never count serialized body text.
                tail_rows = session._execute([
                    'set nativeRows to {}',
                    'repeat with pi from 1 to count paragraphs of boundDoc',
                    'set p to text object of paragraph pi of boundDoc',
                    f'if (content of p as text) is {apple_string(following)} & return then set end of nativeRows to {{start of content of p,end of content of p}}',
                    'end repeat',
                ])
                assert len(tail_rows) == 1 and len(tail_rows[0]) == 2
                tail_start, paragraph_end = tail_rows[0]
                assert type(tail_start) is int and type(paragraph_end) is int
                tail_end = tail_start+len(following.encode('utf-16-le'))//2
                assert paragraph_end == tail_end+1
                case['native_normal_tail'] = _format(session, tail_start, tail_end)
                expected_tail = [[tail_start, tail_end, following, False, False, False]]
                assert case['native_normal_tail'] == expected_tail
                journal.check(label, 'following_normal_text_has_no_notice_style')
                session.save_docx(output/'notices.docx')
                session.export_pdf(output/'notices.pdf')
                case['ooxml'] = inspect_ooxml(output/'notices.docx', display, following,
                                               block=mode != 'inline', table=mode == 'block')
                journal.check(label, 'native_save_pdf_and_persisted_ooxml_style')
                assert next(row for row in inventory(output, 'after-save') if row[0] == sentinel) == sentinel_before
                journal.check(label, 'unsaved_sentinel_hash_and_saved_state_unchanged')
            finally:
                _retain_runtime(session, output/'native-runtime')
        observed = inventory(output, 'after-owned-close')
        assert next(row for row in observed if row[0] == sentinel) == sentinel_before
        assert [row for row in observed if row[0] != sentinel] == starting
        journal.check(label, 'owned_create_context_closed_before_sentinel')
        source_hash = hashlib.sha256((output/'notices.docx').read_bytes()).hexdigest()
        with open_document(output/'notices.docx', engine='msoffice', read_only=False, visible=False) as session:
            try:
                assert _body(session) == body
                assert _format(session, own_range.Start, own_range.Start+extent) == expected_format
                assert _format(session, tail_start, tail_end) == expected_tail
                checkpoint = session.degradation_checkpoint()
                session.add_paragraph('REOPEN-EDIT-' + label)
                assert 'REOPEN-EDIT-' + label in _body(session)[0][0]
                session.rollback_degradation_checkpoint(checkpoint)
                assert _body(session) == body
                journal.check(label, 'reopen_native_style_real_edit_and_rollback')
            finally:
                _retain_runtime(session, output/'reopen-runtime')
        assert hashlib.sha256((output/'notices.docx').read_bytes()).hexdigest() == source_hash
        case['docx_sha256'] = source_hash
        observed = inventory(output, 'after-reopen-close')
        assert next(row for row in observed if row[0] == sentinel) == sentinel_before
        assert [row for row in observed if row[0] != sentinel] == starting
        journal.check(label, 'source_unchanged_and_reopen_context_closed')
        case['sentinel_close'] = close_sentinel(output, sentinel, sentinel_token)
        sentinel = None
        assert inventory(output, 'after-independent-sentinel-close') == starting
        journal.check(label, 'independent_exact_sentinel_close_and_original_inventory')
        import pdfplumber
        with pdfplumber.open(output/'notices.pdf') as pdf:
            pdf_text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        assert pdf_text.count(code) == 1 and following in pdf_text
        assert 'PARTIAL-NOTICE-REMNANT' not in pdf_text
        case['pdf_text'] = pdf_text
        journal.check(label, 'pdf_visible_notice_once_and_normal_tail')
        case['passed'] = True
        journal.flush()
    except BaseException as error:
        case['error'] = {'type': type(error).__name__, 'message': str(error),
                         'quarantined': getattr(session, '_quarantined', False)}
        if sentinel:
            case['retained_sentinel'] = sentinel
        (output/'failure.txt').write_text(traceback.format_exc(), encoding='utf-8')
        _retain_runtime(session, output/'failed-runtime')
        journal.flush()
        raise


def main(output):
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    journal = Journal(output)
    try:
        for mode in ('inline', 'block', 'table-failure'):
            for nonempty in (False, True):
                _case(journal, mode, nonempty)
        for relative, expected in journal.data['source_hashes'].items():
            assert hashlib.sha256((ROOT/relative).read_bytes()).hexdigest() == expected, 'Source changed during native run'
        journal.data['checks']['sources_unchanged_during_run'] = True
        journal.data['passed'] = True
    except BaseException as error:
        journal.data['error'] = {'type': type(error).__name__, 'message': str(error)}
    journal.flush()
    return journal.data


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    if not args.execute:
        parser.error('--execute required')
    report = main(args.output)
    print(json.dumps({'passed': report['passed'], 'error': report.get('error')}))
    sys.exit(0 if report['passed'] else 1)
