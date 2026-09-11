"""Stage one: inactive noncollapsed reservation, one notice, duplicate zero-write.

PREPARATION ONLY until root grants the exclusive Word lease. --execute is also
mandatory. Each run requires a nonexistent output directory. No retry, fallback
fixture repair, UI action, settings change, or full four-method parity claim.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from fixtures.microsoft_parity import macos_word_inline_rule_feasibility as safe
from fixtures.microsoft_parity.macos_word_quality_feasibility import (
    bound_guard, sentinel_guard, seed_commands, snapshot_commands, state_valid,
    sha, source_pair_matches,
)
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage

DICTIONARY = Path('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
SOURCES = [*sorted((ROOT/'skills/WPSComposer').rglob('*.py')),
           *sorted((ROOT/'fixtures/microsoft_parity').glob('*.py')),
           ROOT/'tests/msoffice/test_macos_word_quality_fixture.py', ROOT/'pyproject.toml']
TITLE = 'QUALITY TITLE 中文😀'
MESSAGE = 'Notice 中文😀 retained'
DISPLAY = TITLE+'\r[QUALITY_TEST] '+MESSAGE
ISSUE = {'code': 'QUALITY_TEST', 'message': MESSAGE, 'nodeId': 'fixture-one'}
ANCHOR = 'wpsc_document_quality_anchor'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PENDING = ['add_document_quality_notice first call', 'explicit bookmark first-paragraph End',
           'default bookmark Start', 'invalid/unopened targets zero-write',
           'consecutive upserts', 'fallback/failure matrix', 'UI edit Undo save reopen',
           'M5 numbering/field convergence', 'visual PDF inspection']


def same(a, b):
    return json.dumps(a, ensure_ascii=False, sort_keys=True) == json.dumps(b, ensure_ascii=False, sort_keys=True)


def reservation_valid(rows, path, sentinel, point):
    if not (isinstance(rows, list) and len(rows) == 1 and isinstance(rows[0], list) and len(rows[0]) == 10):
        return False
    r = rows[0]
    return (r[:3] == ['reservation', path, sentinel]
            and all(type(r[i]) is int for i in range(3, 9))
            and r[3:9] == [point-3, point, 2, 2, point, point] and r[9] is True)


def notice_valid(rows, point, before_end, cursor):
    if not (isinstance(rows, list) and len(rows) == 2 and isinstance(rows[0], list)
            and len(rows[0]) == 10 and isinstance(rows[1], list) and len(rows[1]) == 9):
        return False
    n = len(DISPLAY.encode('utf-16-le'))//2
    r, style = rows
    return (all(type(r[i]) is int for i in (1, 2, 3, 4, 5, 7, 8))
            and r == ['notice', point, point+n+2, before_end+n+2, 3, 2, DISPLAY, point, point, True]
            and r[9] is True and type(cursor) is int and cursor == point+n+2
            and same(style, ['style', True, True, True, 0, 3, True, True, False]))


def reservation_commands():
    return ['set qr to text object of selection of boundWindow',
            'set qa to text object of selection of active window',
            f'set qb to bookmark "{ANCHOR}" of boundDoc',
            'set nativeRows to {{"reservation",posix full name of document of selection of boundWindow as text,'
            'name of document of active window as text,start of content of qr,end of content of qr,'
            'start of content of qa,end of content of qa,start of bookmark of qb,end of bookmark of qb,empty of qb}}']


def notice_commands(point):
    # Independent native observations, not the helper's success ACK parser.
    return ['set qt to table 2 of boundDoc',
            'set qr to text object of (get cell from table qt row 1 column 1)',
            f'set qdisplay to create range boundDoc start {point} end ({point}+{len(DISPLAY.encode("utf-16-le"))//2})',
            f'set qb to bookmark "{ANCHOR}" of boundDoc',
            'set nativeRows to {{"notice",start of content of text object of qt,end of content of text object of qt,'
            'end of content of text object of boundDoc,count tables of boundDoc,2,content of qdisplay as text,'
            'start of bookmark of qb,end of bookmark of qb,empty of qb},'
            '{"style",italic of font object of qr,(color of font object of qr is {40092,0,1542}),'
            '(background pattern color of shading of qr is {64764,59624,59110}),'
            'space before of paragraph format of qr,space after of paragraph format of qr,'
            'keep together of paragraph format of qr,(outline level of paragraph format of qr is outline level body text),'
            'allow break across pages of row 1 of qt}}']


def xml_preserved(before, after):
    """Compare every old OOXML element, field and style after removing one box.

    Bookmark geometry is checked natively; revision identifiers are volatile.
    Everything else, including old tables, paragraph formats and field codes,
    must match. This is inspection, never a DOCX mutation backend.
    """
    old, new = ET.fromstring(before), ET.fromstring(after)
    candidates = [e for e in new.iter(W+'tbl') if ''.join(t.text or '' for t in e.iter(W+'t')) == DISPLAY.replace('\r', '')]
    if len(candidates) != 1:
        return False
    for parent in new.iter():
        if candidates[0] in list(parent):
            parent.remove(candidates[0])
            break
    def normalized(node):
        attrs = tuple(sorted((k, v) for k, v in node.attrib.items()
                             if not k.split('}')[-1].startswith('rsid')
                             and k.split('}')[-1] not in ('paraId', 'textId')))
        return (node.tag, attrs, node.text or '', tuple(normalized(c) for c in node
                if c.tag not in (W+'bookmarkStart', W+'bookmarkEnd')))
    return normalized(old) == normalized(new)


def record_failure(output, report, exc, label='primary'):
    key = 'error' if label == 'primary' else label+'_error'
    report.setdefault(key, {'type': type(exc).__name__, 'message': str(exc)})
    report['status'] = 'FAIL'
    try:
        (output/(label+'-failure.txt')).write_text(''.join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except BaseException as write_error:
        report.setdefault('failure_retention_error', str(write_error))


def quarantine(output, report, stage, exc, owner):
    try:
        safe.mark_native_uncertainty(output, report, stage, exc, owner)
    except BaseException as retention_error:
        # Retention is local I/O. Its failure must never mask the originating
        # native failure or permit a normal close/final inventory afterward.
        report['quarantined'] = True
        if owner:
            owner._quarantined = True
        record_failure(output, report, retention_error, 'quarantine_retention')


def cleanup(session, output, report):
    """One cleanup attempt; uncertainty forbids every later native call."""
    if report.get('cleanup_finished'):
        return
    report['cleanup_finished'] = True
    report['quarantined'] = bool(report.get('native_uncertainty') or (session and session._quarantined))
    name = report.get('sentinel_name')
    if (not report['quarantined'] and session and name
            and not report.get('sentinel_closed') and not report.get('sentinel_cleanup_attempted')):
        try:
            safe.close_sentinel_after_owned(session, output, report, name, report.get('sentinel_token', ''))
        except BaseException as exc:
            record_failure(output, report, exc, 'cleanup')
    report['quarantined'] = bool(report.get('native_uncertainty') or (session and session._quarantined))
    report['remaining_sentinel'] = None if report.get('sentinel_closed') else name
    if not report['quarantined'] and 'inventory_before' in report:
        try:
            report['inventory_final'] = safe.independent_inventory(output, 'final', report, session)
            report['checks']['inventory_preserved'] = same(report['inventory_final'], report['inventory_before'])
        except BaseException as exc:
            record_failure(output, report, exc, 'inventory')
    report['quarantined'] = bool(report.get('native_uncertainty') or (session and session._quarantined))


def pdf_checks(output):
    import fitz
    with fitz.open(output/'after.pdf') as pdf:
        text = '\n'.join(page.get_text() for page in pdf)
        for index, page in enumerate(pdf):
            page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(output/f'page-{index+1}.png')
    (output/'pdf-text.txt').write_text(text)
    return text.count('QUALITY TITLE') == text.count('[QUALITY_TEST]') == 1 and 'PREFIX' in text and 'TAIL' in text


def run(output, *, execute=False):
    if execute is not True:
        raise ValueError('Explicit execute=True is required; root must separately grant the Word lease')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'status': 'FAIL', 'scope': 'stage-one reservation/first upsert/duplicate only',
              'pending': PENDING.copy(), 'checks': {}, 'steps': [], 'confirmed': []}
    session = reopened = None

    def flush():
        path = output/'report.pending'
        path.write_text(json.dumps(report, ensure_ascii=False, indent=2))
        path.replace(output/'report.json')

    def observe(label, commands, validator=None, *, mutate=False, owner=None):
        owner = owner or session
        report['current_step'] = label
        flush()
        rows = (owner._execute_structural if mutate else owner._execute)(bound_guard(owner._bound_path)+commands)
        report['steps'].append({'label': label, 'rows': rows})
        flush()
        if validator is not None and not validator(rows):
            quarantine(output, report, label, 'Native observation invalid', owner)
            raise AssertionError(label+' native observation invalid')
        return rows

    try:
        report['source_hashes'] = retain_sources(output, SOURCES)
        report['dictionary_sha256'] = sha(DICTIONARY)
        shutil.copy2(DICTIONARY, output/'localWord.sdef')
        report['inventory_before'] = safe.independent_inventory(output, 'before', report)
        report['current_step'] = 'open-owned'; flush()
        session = MacWordSession.new_document(visible=False)
        session._retain_evidence = True
        observe('seed', seed_commands('middle-table-recovery'), lambda r: same(r, [['seed', True]]), mutate=True)
        before = observe('before-snapshot', snapshot_commands(), state_valid)
        if before[0][4:6] != [2, 2]:
            raise AssertionError('Expected two existing tables and fields')
        session.save_docx(output/'before.docx')
        report['before_sha256'] = sha(output/'before.docx')
        token = 'QUALITY STAGE ONE SENTINEL 中文😀 '+uuid4().hex
        report['sentinel_token'] = token
        rows = observe('sentinel-create', ['set qualitySentinel to make new document',
            f'set content of text object of qualitySentinel to {apple_string(token)}',
            'set nativeRows to {{name of qualitySentinel as text,version as text}}'],
            lambda r: isinstance(r, list) and len(r) == 1 and isinstance(r[0], list) and len(r[0]) == 2
                      and all(isinstance(v, str) and v for v in r[0]))
        name, report['word_version'] = rows[0]
        report['sentinel_name'] = name
        report['sentinel_before'] = sentinel_preimage(safe.independent_inventory(output, 'sentinel-before', report, session), name, token)
        point = next(r[2] for r in before if r[0:2] == ['bookmark', 'wpsc_quality_splice'])
        report['point'] = point
        observe('inactive-noncollapsed-selection', [
            'activate object boundWindow',
            f'set selection start of selection of boundWindow to {point-3}',
            f'set selection end of selection of boundWindow to {point}',
            *sentinel_guard(name, token), 'activate object (active window of qualitySentinel)',
            'set selection start of selection of active window to 2',
            'set selection end of selection of active window to 2', 'set nativeRows to {{"selected",true}}'],
            lambda r: same(r, [['selected', True]]))
        report['current_step'] = 'reserve-public'; flush()
        if session.reserve_document_quality_anchor(title=TITLE) is not None:
            raise AssertionError('Reservation return differs')
        observe('reservation', reservation_commands(), lambda r: reservation_valid(r, session._bound_path, name, point))
        empty = observe('empty-snapshot', snapshot_commands(), state_valid)
        projected = [r.copy() for r in empty if r[0:2] != ['bookmark', ANCHOR]]
        projected[0][6] -= 1
        # Removing the added bookmark preserves all original ordinals and hashes.
        if not same(projected, before) or session._quality_notice_anchor_position != point:
            raise AssertionError('Empty reservation changed document or cursor')
        report['confirmed'].append('reserved'); report['checks']['empty_reservation'] = True; flush()
        report['current_step'] = 'first-upsert-public'; flush()
        if session.upsert_document_quality_notice(ISSUE.copy()) is not None:
            raise AssertionError('Upsert return differs')
        cursor = session._quality_notice_anchor_position
        observe('first-notice', notice_commands(point), lambda r: notice_valid(r, point, before[0][1], cursor))
        after = observe('after-snapshot', snapshot_commands(), state_valid)
        report['confirmed'].append('first-notice'); report['checks']['first_notice'] = True; flush()
        # Preserve confirmed first-item artifacts before testing duplicate behavior.
        session.save_docx(output/'partial.docx')
        session.export_pdf(output/'partial.pdf')
        scripts = {str(p): sha(p) for p in session.staging_root.glob('*.applescript')}
        report['current_step'] = 'duplicate-public'; flush()
        if session.upsert_document_quality_notice(dict(ISSUE, message='must not replace')) is not None:
            raise AssertionError('Duplicate return differs')
        report['checks']['duplicate_zero_transport'] = scripts == {str(p): sha(p) for p in session.staging_root.glob('*.applescript')}
        duplicate = observe('duplicate-snapshot', snapshot_commands(), state_valid)
        report['checks']['duplicate_unchanged'] = same(duplicate, after) and session._quality_notice_anchor_position == cursor
        if not all(report['checks'][key] for key in ('duplicate_zero_transport', 'duplicate_unchanged')):
            raise AssertionError('Duplicate performed work')
        report['confirmed'].append('duplicate'); flush()
        session.save_docx(output/'after.docx'); session.export_pdf(output/'after.pdf')
        report['checks']['sentinel_unchanged'] = same(report['sentinel_before'], sentinel_preimage(
            safe.independent_inventory(output, 'sentinel-after', report, session), name, token))
        session.close()
        # Close sentinel before reopening, using only the reviewed independent guard.
        safe.close_sentinel_after_owned(session, output, report, name, token)
        digest = sha(output/'after.docx')
        report['current_step'] = 'open-readonly'
        report['reopen_requested_path'] = str(output/'after.docx'); flush()
        reopened = MacWordSession.open_document(output/'after.docx', read_only=True, visible=False)
        reopened._retain_evidence = True
        again = observe('reopen-snapshot', snapshot_commands(), state_valid, owner=reopened)
        observe('reopen-notice', notice_commands(point), lambda r: notice_valid(r, point, before[0][1], cursor), owner=reopened)
        report['checks']['reopen_preserved'] = same(again, after)
        reopened.close()
        report['checks']['reopen_closed_unchanged'] = reopened._closed and not reopened._quarantined and sha(output/'after.docx') == digest
        report['checks']['xml_preserved'] = xml_preserved(safe.retain_xml(output, 'before'), safe.retain_xml(output, 'after'))
        report['checks']['styles_preserved'] = (output/'before-xml/word/styles.xml').read_bytes() == (output/'after-xml/word/styles.xml').read_bytes()
        report['checks']['pdf_text_once'] = pdf_checks(output)
    except BaseException as exc:
        record_failure(output, report, exc)
        owner = reopened or session
        if report.get('current_step') in ('open-owned', 'open-readonly'):
            # A creation/open factory can submit before returning an owner. No
            # returned session is not proof that Word received no AppleEvent.
            quarantine(output, report, report['current_step'], exc, owner)
        elif owner and not owner._closed:
            # Preserve confirmed partial files and live identity; never repair or
            # issue a new observation after a failed public/native operation.
            quarantine(output, report, report.get('current_step', 'run'), exc, owner)
    finally:
        if reopened and reopened._quarantined:
            report['quarantined'] = True
        cleanup(reopened or session, output, report)
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner:
                report[label+'-identity'] = {'path': owner._bound_path, 'staging_root': str(owner.staging_root), 'closed': owner._closed, 'quarantined': owner._quarantined}
                try:
                    if owner.staging_root.exists():
                        shutil.copytree(owner.staging_root, output/label)
                except BaseException as exc:
                    record_failure(output, report, exc, label+'-retention')
        try:
            if 'source_hashes' in report:
                report['checks']['sources_unchanged'] = all(source_pair_matches(ROOT/p, output/'source'/p, value) for p, value in report['source_hashes'].items())
            if 'dictionary_sha256' in report:
                report['checks']['dictionary_unchanged'] = source_pair_matches(DICTIONARY, output/'localWord.sdef', report['dictionary_sha256'])
            if 'before_sha256' in report:
                report['checks']['before_file_unchanged'] = sha(output/'before.docx') == report['before_sha256']
            for label in ('before', 'after', 'partial'):
                if (output/(label+'.docx')).exists():
                    safe.retain_xml(output, label)
            report['artifact_hashes'] = {str(p.relative_to(output)): sha(p) for p in output.rglob('*')
                if p.is_file() and 'source' not in p.relative_to(output).parts
                and p.name not in ('report.json', 'report.pending')}
        except BaseException as exc:
            record_failure(output, report, exc, 'evidence')
        required = ('empty_reservation', 'first_notice', 'duplicate_zero_transport', 'duplicate_unchanged',
                    'sentinel_unchanged', 'owned_closed_with_sentinel_preserved', 'reopen_preserved',
                    'reopen_closed_unchanged', 'xml_preserved', 'styles_preserved', 'pdf_text_once',
                    'sources_unchanged', 'dictionary_unchanged', 'before_file_unchanged', 'inventory_preserved')
        if (all(report['checks'].get(k) is True for k in required) and not report.get('quarantined')
                and not report.get('remaining_sentinel') and not any(k == 'error' or k.endswith(('_error', '_failure')) for k in report)):
            report['status'] = 'PASS_STAGE_ONE_AUTOMATED'
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Requires --execute and a separate root-granted exclusive Word lease', file=sys.stderr)
        return 2
    return 0 if run(args.output, execute=True)['status'] == 'PASS_STAGE_ONE_AUTOMATED' else 1


if __name__ == '__main__':
    raise SystemExit(main())
