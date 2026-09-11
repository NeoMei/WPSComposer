"""One guarded inline constructor hypothesis; no production or WordArt runner.

Requires an exclusive root-granted Word lease and explicit --execute. The sole
constructor is imported unchanged from the compile-only document-container
candidate. An ordinary failure is retained, never repaired or retried.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, _JSON
from fixtures.microsoft_parity.macos_word_inline_document_variant import build_commands
from fixtures.microsoft_parity.macos_word_rules_probe import readback_commands, verify_native_rows
from fixtures.microsoft_parity.macos_word_quality_feasibility import bound_guard, sentinel_guard, sha, source_pair_matches
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_paragraph_rule import sentinel_preimage

DICTIONARY = Path('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
PREFIX = 'INLINE RULE PROBE 中文😀'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
V = '{urn:schemas-microsoft-com:vml}'
O = '{urn:schemas-microsoft-com:office:office}'
# Include transitive fixture helpers as well as all shipped Python dependencies.
SOURCES = [*sorted((ROOT/'skills/WPSComposer').rglob('*.py')),
           *sorted((ROOT/'fixtures/microsoft_parity').glob('*.py')),
           ROOT/'tests/msoffice/test_macos_word_inline_rule_feasibility.py', ROOT/'pyproject.toml']


def native_valid(rows):
    return verify_native_rows('inline', rows)


def inline_xml_valid(payload):
    root = ET.fromstring(payload)
    rules = [e for e in root.iter(V+'rect') if e.get(O+'hr') in ('t', 'true')]
    inline = [e for pict in root.iter(W+'pict') for e in pict.iter(V+'rect') if e in rules]
    return (len(rules) == len(inline) == 1
            and not any(True for tag in (W+'pBdr', W+'drawing', V+'imagedata') for _ in root.iter(tag)))


def retain_xml(output, label):
    with ZipFile(output/(label+'.docx')) as package:
        # Keep every XML part verbatim, including relationships and settings.
        for member in package.namelist():
            if member.endswith(('.xml', '.rels')):
                target = output/(label+'-xml')/member
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(package.read(member))
        return package.read('word/document.xml')


def pdf_rule_matches(drawing, page_rect, dimensions):
    def number(value):
        return type(value) in (int, float) and math.isfinite(value)

    if not (isinstance(dimensions, (list, tuple)) and len(dimensions) == 2
            and all(number(v) and v > 0 for v in dimensions)):
        return False
    rect = drawing.get('rect')
    items = drawing.get('items', [])
    if rect is None or len(items) != 1 or len(items[0]) != 3:
        return False
    item = items[0]
    if drawing.get('type') == 'f' and item[0] == 're':
        thickness = rect.height
        color = drawing.get('fill')
        if drawing.get('fill_opacity') != 1 or tuple(item[1]) != tuple(rect):
            return False
    elif drawing.get('type') == 's' and item[0] == 'l':
        start, end = item[1:]
        thickness = drawing.get('width')
        color = drawing.get('color')
        if (drawing.get('stroke_opacity') != 1 or drawing.get('dashes') not in (None, '[] 0')
                or abs(start.y-end.y) > 0.05 or rect.height > 0.05
                or abs(abs(start.x-end.x)-rect.width) > 0.05):
            return False
    else:
        return False
    return (isinstance(color, (list, tuple)) and len(color) == 3
            and all(number(v) and 0 <= v <= 1 for v in color)
            and number(thickness) and 0 < thickness < 10
            and all(number(v) for v in rect)
            and rect.width > 10*max(rect.height, thickness)
            and abs(rect.width-dimensions[0]) <= 2 and abs(thickness-dimensions[1]) <= 2
            and page_rect.x0 <= rect.x0 <= rect.x1 <= page_rect.x1
            and page_rect.y0 <= rect.y0 <= rect.y1 <= page_rect.y1)


def pdf_checks(output, rows):
    import fitz
    dimensions = rows[0][3:] if native_valid(rows) else None
    visible = False
    observations = []
    with fitz.open(output/'after.pdf') as pdf:
        text = '\n'.join(page.get_text() for page in pdf)
        for index, page in enumerate(pdf):
            page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(output/f'page-{index+1}.png')
            for drawing in page.get_drawings():
                observations.append({'page': index+1, 'drawing': drawing})
                # Compare a thin opaque vector's geometry to native dimensions;
                # arbitrary drawing count, images and visible text cannot pass.
                if pdf_rule_matches(drawing, page.rect, dimensions):
                    visible = True
    (output/'pdf-text.txt').write_text(text)
    (output/'pdf-drawings.json').write_text(json.dumps(observations, default=str, indent=2))
    return {'pdf_anchor_visible': 'INLINE RULE PROBE' in text, 'pdf_native_geometry_visible': visible}


def retain_transport_log(path, report, *parts):
    def text(value):
        return value.decode('utf-8', errors='replace') if isinstance(value, bytes) else str(value or '')
    try:
        path.write_text('\n'.join(text(part) for part in parts))
    except BaseException:
        report['independent_log_failure'] = traceback.format_exc()


def mark_native_uncertainty(output, report, stage, error, session=None):
    """Quarantine the runner without rewriting a verified owned-close fact."""
    if report.get('native_uncertainty'):
        return
    report['status'] = 'FAIL'
    report['quarantined'] = True
    report['sentinel_state'] = 'closed-acknowledged' if report.get('sentinel_closed') else 'unknown'
    marker = {'schema': 1, 'stage': stage, 'error': str(error),
              'action': 'Stop native calls; root must independently reconcile this task before another native run.',
              'owned_document': {'path': getattr(session, '_bound_path', None),
                                 'closed': bool(session and session._closed),
                                 'staging_root': str(session.staging_root) if session else None},
              'sentinel': {'name': report.get('sentinel_name'), 'preimage': report.get('sentinel_before'),
                           'state': report['sentinel_state'], 'close_script': str(output/'sentinel-close.applescript')}}
    report['native_uncertainty'] = marker
    report['recovery_marker'] = str(output/'native-uncertainty.json')
    try:
        (output/'native-uncertainty.json').write_text(json.dumps(marker, ensure_ascii=False, indent=2))
    except BaseException:
        report['recovery_marker_failure'] = traceback.format_exc()
    # An independent observation can also fail while the owned context is live.
    # Its normal __exit__ must not send a close event into the uncertain host.
    # Already verified closed sessions keep their real closed/quarantine flags.
    if session and not session._closed and not session._quarantined:
        session._retain('Independent inline probe transport completion uncertain')


def independent_inventory(output, label, report, session=None):
    if report.get('native_uncertainty'):
        raise RuntimeError('Runner native completion uncertain; inventory refused')
    try:
        rows = inventory(output, label)
        if not (isinstance(rows, list) and all(isinstance(row, list) and len(row) == 4
                and isinstance(row[0], str) and isinstance(row[1], str) and type(row[2]) is bool
                and isinstance(row[3], str) and re.fullmatch('[0-9a-f]{64}', row[3]) for row in rows)):
            raise RuntimeError('Independent inventory acknowledgement invalid')
        return rows
    except BaseException as exc:
        # The shared inventory helper hides its submission boundary. Its errors
        # cannot prove no AppleEvent was issued, so fail closed locally.
        mark_native_uncertainty(output, report, label, exc, session)
        retain_transport_log(output/(label+'-uncertainty.log'), report, type(exc).__name__,
                             str(exc), getattr(exc, 'stdout', None), getattr(exc, 'stderr', None))
        raise


def close_sentinel_after_owned(session, output, report, name, token):
    report['sentinel_cleanup_attempted'] = True
    if report.get('native_uncertainty'):
        raise RuntimeError('Runner native completion uncertain; sentinel retained for reconciliation')
    if not session._closed or session._quarantined:
        raise RuntimeError('Owned close unverified; sentinel retained')
    observed = independent_inventory(output, 'after-owned-close', report, session)
    report['inventory_after_owned_close'] = observed
    if ([r for r in observed if r[0] == name] != [report['sentinel_before']]
            or [r for r in observed if r[0] != name] != report['inventory_before']):
        report['sentinel_close_state'] = 'identity-refused-before-close'
        raise RuntimeError('Inventory or sentinel changed; retained')
    report['checks']['owned_closed_with_sentinel_preserved'] = True
    script = output/'sentinel-close.applescript'
    report['sentinel_close_state'] = 'not-submitted'
    script.write_text(_JSON+'\nif not application "Microsoft Word" is running then error "SENTINEL_UNAVAILABLE"\n'
        +'tell application "Microsoft Word"\n'+'\n'.join(sentinel_guard(name, token)
        +['close qualitySentinel saving no', 'set nativeRows to {{"sentinel-closed"}}'])
        +'\nend tell\nreturn my jsonRows(nativeRows)\n')
    try:
        result = subprocess.run(['/usr/bin/osascript', str(script)], capture_output=True, text=True, timeout=30)
    except (FileNotFoundError, PermissionError) as exc:
        # Popen could not execute the child; no close command was submitted.
        retain_transport_log(script.with_suffix('.log'), report, type(exc).__name__, str(exc))
        raise
    except BaseException as exc:
        report['sentinel_close_state'] = 'unknown'
        mark_native_uncertainty(output, report, 'sentinel-close', exc, session)
        retain_transport_log(script.with_suffix('.log'), report, type(exc).__name__, str(exc),
                             getattr(exc, 'stdout', None), getattr(exc, 'stderr', None))
        raise
    try:
        acknowledged = result.returncode == 0 and json.loads(result.stdout) == [['sentinel-closed']]
    except (ValueError, TypeError):
        acknowledged = False
    # These explicit guard errors precede the close command in this exact script.
    refused = result.returncode != 0 and re.search(
        r'\b(?:QUALITY_SENTINEL_NAME|QUALITY_SENTINEL_TEXT|QUALITY_SENTINEL_SAVED|SENTINEL_UNAVAILABLE) \(-2700\)\s*$', result.stderr)
    if not acknowledged and not refused:
        report['sentinel_close_state'] = 'unknown'
        mark_native_uncertainty(output, report, 'sentinel-close',
                                'Unverified close ACK: '+result.stdout+'\n'+result.stderr, session)
    elif refused:
        report['sentinel_close_state'] = 'identity-refused-before-close'
    else:
        report['sentinel_close_state'] = 'closed-acknowledged'
        report['sentinel_closed'] = True
    retain_transport_log(script.with_suffix('.log'), report, result.stdout, result.stderr)
    if not acknowledged:
        raise RuntimeError('Sentinel close acknowledgement invalid')


def retain_partial(session, output, report):
    """Capture a single failed state without changing its semantic verdict."""
    if session._quarantined or session._closed:
        report['partial_retention_skipped'] = 'session closed or quarantined'
        return
    try:
        rows = session._execute(bound_guard(session._bound_path)+[
            'set nativeRows to {{"partial-counts",count inline shapes of boundDoc,count shapes of boundDoc}}'])
        report['steps'].append({'label': 'partial-counts', 'rows': rows})
        session.save_docx(output/'partial.docx')
        session.export_pdf(output/'partial.pdf')
    except BaseException:
        report['partial_retention_error'] = traceback.format_exc()


def run(output, *, execute=False):
    if execute is not True:
        raise ValueError('Explicit execute=True is required')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'status': 'FAIL', 'checks': {}, 'steps': [],
              'scope': 'one native constructor hypothesis only; public method/UI/parity remain pending'}
    session = reopened = None
    name = token = None

    def flush():
        pending = output/'report.pending'
        pending.write_text(json.dumps(report, ensure_ascii=False, indent=2))
        pending.replace(output/'report.json')

    def step(label, commands, *, mutate=False):
        report['current_step'] = label
        flush()
        action = session._execute_structural if mutate else session._execute
        rows = action(bound_guard(session._bound_path)+commands)
        report['steps'].append({'label': label, 'rows': rows})
        flush()
        return rows

    try:
        report['source_hashes'] = retain_sources(output, SOURCES)
        report['dictionary_sha256'] = sha(DICTIONARY)
        shutil.copy2(DICTIONARY, output/'localWord.sdef')
        report['inventory_before'] = independent_inventory(output, 'before', report)
        with MacWordSession.new_document(visible=False) as session:
            session._retain_evidence = True
            token = 'INLINE RULE SENTINEL 中文😀 '+uuid4().hex
            rows = step('sentinel-create', ['set qualitySentinel to make new document',
                f'set content of text object of qualitySentinel to {apple_string(token)}',
                'set nativeRows to {{name of qualitySentinel as text,version as text}}'])
            if not (len(rows) == 1 and len(rows[0]) == 2 and isinstance(rows[0][0], str) and rows[0][0]):
                raise RuntimeError('Sentinel creation acknowledgement invalid')
            name, report['word_version'] = rows[0]
            report['sentinel_name'] = name
            report['sentinel_before'] = sentinel_preimage(independent_inventory(output, 'sentinel-before', report, session), name, token)
            rows = step('seed', [f'set content of text object of boundDoc to {apple_string(PREFIX)} & return',
                'set nativeRows to {{"seed",count inline shapes of boundDoc,count shapes of boundDoc}}'], mutate=True)
            if json.dumps(rows) != json.dumps([['seed', 0, 0]]):
                raise AssertionError('Seed native object counts differ')
            session.save_docx(output/'before.docx')
            report['before_sha256'] = sha(output/'before.docx')
            try:
                rows = step('inline-document-container-only', build_commands(), mutate=True)
                report['native_rows'] = rows
                report['checks']['native_readback'] = native_valid(rows)
                if not report['checks']['native_readback']:
                    raise AssertionError('Native count/type/geometry mismatch')
            except BaseException as exc:
                report['constructor_error'] = {'type': type(exc).__name__, 'message': str(exc)}
                (output/'constructor-failure.txt').write_text(traceback.format_exc())
                # Ordinary acknowledged errors may leave a changed document.
                # Uncertain completion is quarantined by the shared transport.
                retain_partial(session, output, report)
                flush()
                raise
            session.save_docx(output/'after.docx')
            session.export_pdf(output/'after.pdf')
            report['sentinel_after'] = sentinel_preimage(independent_inventory(output, 'sentinel-after', report, session), name, token)
            report['checks']['sentinel_unchanged'] = report['sentinel_after'] == report['sentinel_before']
        close_sentinel_after_owned(session, output, report, name, token)
        name = None
        digest = sha(output/'after.docx')
        with MacWordSession.open_document(output/'after.docx', read_only=True, visible=False) as reopened:
            reopened._retain_evidence = True
            report['reopen_rows'] = reopened._execute(bound_guard(reopened._bound_path)+readback_commands('inline'))
            report['checks']['reopen_native_readback'] = native_valid(report['reopen_rows'])
            report['checks']['reopen_matches_creation'] = json.dumps(report['reopen_rows']) == json.dumps(report['native_rows'])
        report['checks']['reopen_closed'] = reopened._closed and not reopened._quarantined
        report['checks']['reopen_source_unchanged'] = sha(output/'after.docx') == digest
        before_xml = retain_xml(output, 'before')
        after_xml = retain_xml(output, 'after')
        report['checks']['persisted_native_inline_xml'] = inline_xml_valid(after_xml)
        text = lambda payload: ''.join(n.text or '' for n in ET.fromstring(payload).iter(W+'t'))
        report['checks']['prefix_text_preserved'] = text(before_xml) == text(after_xml) == PREFIX
        report['checks'].update(pdf_checks(output, report['native_rows']))
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if session and name and not report.get('sentinel_cleanup_attempted'):
            try:
                close_sentinel_after_owned(session, output, report, name, token)
                name = None
            except BaseException:
                report['cleanup_failure'] = traceback.format_exc()
        if report.get('sentinel_closed'):
            name = None
        report['remaining_sentinel'] = name
        report['quarantined'] = bool(report.get('native_uncertainty') or (session and session._quarantined) or (reopened and reopened._quarantined))
        for owner, label in ((session, 'native-runtime'), (reopened, 'reopen-runtime')):
            if owner:
                report[label+'-identity'] = {'path': owner._bound_path, 'staging_root': str(owner.staging_root), 'closed': owner._closed}
                try:
                    if owner.staging_root and owner.staging_root.exists():
                        shutil.copytree(owner.staging_root, output/label, dirs_exist_ok=True)
                except BaseException:
                    report[label+'-retention-error'] = traceback.format_exc()
        for label in ('before', 'after', 'partial'):
            if (output/(label+'.docx')).exists():
                try:
                    retain_xml(output, label)
                except BaseException:
                    report[label+'-xml-error'] = traceback.format_exc()
        try:
            if 'source_hashes' in report:
                report['checks']['sources_unchanged'] = all(source_pair_matches(ROOT/p, output/'source'/p, value)
                    for p, value in report['source_hashes'].items())
            if 'dictionary_sha256' in report:
                report['checks']['dictionary_unchanged'] = source_pair_matches(DICTIONARY, output/'localWord.sdef', report['dictionary_sha256'])
            if 'before_sha256' in report:
                report['checks']['before_file_unchanged'] = sha(output/'before.docx') == report['before_sha256']
        except BaseException:
            report['evidence_hash_failure'] = traceback.format_exc()
        if not report['quarantined'] and 'inventory_before' in report:
            try:
                report['inventory_final'] = independent_inventory(output, 'final', report, reopened or session)
                report['checks']['inventory_preserved'] = report['inventory_final'] == report.get('inventory_before')
            except BaseException:
                report['inventory_failure'] = traceback.format_exc()
        report['artifact_hashes'] = {str(p.relative_to(output)): sha(p) for p in output.rglob('*')
            if p.is_file() and p.suffix in ('.docx', '.pdf', '.png', '.xml', '.rels', '.sdef') and 'source' not in p.relative_to(output).parts}
        required = ('native_readback', 'owned_closed_with_sentinel_preserved', 'sentinel_unchanged',
                    'reopen_native_readback', 'reopen_matches_creation', 'reopen_closed', 'reopen_source_unchanged',
                    'persisted_native_inline_xml', 'prefix_text_preserved', 'pdf_anchor_visible', 'pdf_native_geometry_visible',
                    'sources_unchanged', 'dictionary_unchanged', 'before_file_unchanged', 'inventory_preserved')
        if (all(report['checks'].get(k) is True for k in required) and not name and not report['quarantined']
                and not any(k.endswith(('error', 'failure')) for k in report)):
            report['status'] = 'PASS'
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Native inline feasibility requires --execute and the root Word lease', file=sys.stderr)
        return 2
    return 0 if run(args.output, execute=True)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
