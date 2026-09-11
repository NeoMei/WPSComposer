"""Guarded ordinary-error rollback probe; intentionally retains quarantine for UI cleanup."""
from __future__ import annotations

import argparse
from contextlib import contextmanager
import hashlib
import json
import re
from pathlib import Path
import shutil
import subprocess
import sys
import traceback
from uuid import uuid4

ROOT = next(parent for parent in Path(__file__).resolve().parents
            if (parent/'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py').is_file())
sys.path.insert(0, str(ROOT))
from skills.WPSComposer.scripts.msoffice import macos_word_recovery as recovery
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, NativeWordError, _JSON
from skills.WPSComposer.scripts.writer import NativeWriterObjectError
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_recovery import inventory
from fixtures.microsoft_parity.evidence_gate import source_digest

INJECTED_ERROR = 'error "WPSC_INJECTED_ROLLBACK_FAILURE"'
BOOKMARK = 'WPSC_RollbackFailureTarget'
TABLE_TEXT = 'TABLE-SURVIVES-ROLLBACK-ERROR'
SOURCES = [Path(__file__), ROOT/'fixtures/microsoft_parity/macos_word_fields.py',
           ROOT/'fixtures/microsoft_parity/macos_word_recovery.py',
           ROOT/'fixtures/microsoft_parity/evidence_gate.py',
           *[ROOT/'skills/WPSComposer/scripts'/name for name in (
               'msoffice/macos_word_recovery.py', 'msoffice/macos_word_session.py',
               'msoffice/macos_word_fields.py', 'msoffice/macos_runtime.py',
               'msoffice/macos_script.py', 'writer.py')]]


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def inject_after_field_delete(lines):
    source = list(lines)
    points = [index for index, line in enumerate(source) if line == 'delete recoveryField']
    if len(points) != 1 or any('WPSC_INJECTED_ROLLBACK_FAILURE' in line for line in source):
        raise ValueError('Exactly one original field deletion is required for fault injection')
    index = points[0] + 1
    return source[:index] + [INJECTED_ERROR] + source[index:]


@contextmanager
def injected_compiler():
    original = recovery.rollback_commands
    records = []
    def compile_failure(*args, **kwargs):
        source = original(*args, **kwargs)
        injected = inject_after_field_delete(source)
        records.append({'original_sha256': hashlib.sha256('\n'.join(source).encode()).hexdigest(),
                        'injected_sha256': hashlib.sha256('\n'.join(injected).encode()).hexdigest(),
                        'injection_index': injected.index(INJECTED_ERROR), 'injection_count': 1})
        return injected
    recovery.rollback_commands = compile_failure
    try:
        yield records
    finally:
        recovery.rollback_commands = original


def snapshot_source(bound_path, window_id, checkpoint):
    """A separate read-only AppleEvent with exact path/window guards; no session bypass."""
    if (not Path(bound_path).is_absolute() or type(window_id) is not int or window_id < 1
            or type(checkpoint) is not int or checkpoint < 0):
        raise ValueError('Exact owned path, window ID and checkpoint are required')
    lines = ['if not application "Microsoft Word" is running then error "WPSC_DIAGNOSTIC_HOST_ABSENT"',
             'tell application "Microsoft Word"',
             f'set guardedDoc to document {apple_string(Path(bound_path).name)}',
             f'if not ((current application\'s NSString\'s stringWithString:(posix full name of guardedDoc as text))\'s isEqualToString:{apple_string(bound_path)}) then error "WPSC_DIAGNOSTIC_BINDING_CHANGED"',
             f'if (id of active window of guardedDoc) is not {window_id} then error "WPSC_DIAGNOSTIC_WINDOW_CHANGED"',
             *recovery.hash_commands('content of text object of guardedDoc as text', 'bodyHash'),
             f'set prefixRange to create range guardedDoc start 0 end {checkpoint}',
             *recovery.hash_commands('content of prefixRange as text', 'prefixHash'),
             'set nativeRows to {{"owned",posix full name of guardedDoc as text,id of active window of guardedDoc,saved of guardedDoc,count fields of guardedDoc,count tables of guardedDoc,bodyHash,prefixHash}}',
             'repeat with ti from 1 to count tables of guardedDoc',
             'set diagnosticTable to table ti of guardedDoc',
             'set tableText to content of text object of diagnosticTable as text',
             *recovery.hash_commands('tableText', 'tableHash'),
             f'set markerFound to ((current application\'s NSString\'s stringWithString:tableText)\'s containsString:{apple_string(TABLE_TEXT)}) as boolean',
             'set end of nativeRows to {"table",ti as integer,start of content of text object of diagnosticTable,end of content of text object of diagnosticTable,tableHash,markerFound}',
             'end repeat', 'end tell', 'return my jsonRows(nativeRows)']
    return _JSON + '\n'.join(lines)


def validate_snapshot(rows, bound_path, window_id):
    def digest(value):
        return isinstance(value, str) and re.fullmatch('[0-9a-f]{64}', value) is not None
    if (not isinstance(rows, list) or not rows or not isinstance(rows[0], list)
            or len(rows[0]) != 8 or rows[0][:3] != ['owned', bound_path, window_id]
            or type(rows[0][2]) is not int or type(rows[0][3]) is not bool
            or any(type(value) is not int or value < 0 for value in rows[0][4:6])
            or not all(digest(value) for value in rows[0][6:])
            or len(rows) != rows[0][5] + 1):
        raise ValueError('Independent owned-document acknowledgement is invalid')
    for ordinal, row in enumerate(rows[1:], 1):
        if (not isinstance(row, list) or len(row) != 6 or row[:2] != ['table', ordinal]
                or any(type(value) is not int for value in row[1:4])
                or not 0 <= row[2] < row[3] or not digest(row[4]) or type(row[5]) is not bool):
            raise ValueError('Independent table acknowledgement is invalid')
    return rows


def independent_snapshot(output, label, bound_path, window_id, checkpoint):
    path = output/(label+'.applescript')
    path.write_text(snapshot_source(bound_path, window_id, checkpoint), encoding='utf-8')
    try:
        result = subprocess.run(['/usr/bin/osascript', str(path)], capture_output=True, text=True, timeout=30)
    except BaseException as error:
        path.with_suffix('.log').write_text(repr(error), encoding='utf-8')
        raise
    path.with_suffix('.log').write_text(result.stdout+'\n'+result.stderr, encoding='utf-8')
    if result.returncode:
        raise RuntimeError('Independent exact-bound readback failed; owned document retained')
    return validate_snapshot(json.loads(result.stdout), bound_path, window_id)


def script_state(session):
    return {p.name: sha(p) for p in sorted(session.staging_root.glob('*.applescript'))}


def tracking_state(session):
    return (tuple(getattr(session, '_tracked_indexes', ())),
            tuple(getattr(session, '_tracked_references', ())))


def run(output):
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'status':'FAIL_RETAINED', 'checks':{}, 'source_hashes':retain_sources(output, SOURCES),
              'source_digest_before':source_digest(ROOT), 'platform':'darwin', 'engine':'msoffice',
              'scope':'Injected ordinary error after real field deletion; manual UI discard and guarded recovery required'}
    session = None
    def flush():
        pending = output/'report.pending'
        pending.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
        pending.replace(output/'report.json')
    def check(name, condition):
        report['checks'][name] = bool(condition)
        flush()
        if not condition:
            raise AssertionError(name)
    flush()
    try:
        report['inventory_before'] = inventory(output, 'inventory-before'); flush()
        session = MacWordSession.new_document(visible=False)
        session._retain_evidence = True
        report['owned_binding'] = {'path':session._bound_path, 'window_id':session._window_id,
                                   'staging_root':str(session.staging_root),
                                   'quarantine_path':str(session.lock.quarantine_path)}
        flush()
        sentinel_token = 'Rollback failure sentinel 中文😀 '+uuid4().hex
        report['sentinel_token'] = sentinel_token; flush()
        created = session._execute(['set sentinelDoc to make new document',
                                    f'set content of text object of sentinelDoc to {apple_string(sentinel_token)}',
                                    'set nativeRows to {{name of sentinelDoc as text}}'])
        if not isinstance(created, list) or len(created) != 1 or len(created[0]) != 1:
            raise RuntimeError('Sentinel creation identity unavailable')
        sentinel_name = created[0][0]
        report['sentinel_name'] = sentinel_name; flush()
        before = inventory(output, 'sentinel-before')
        matches = [row for row in before if row[0] == sentinel_name]
        check('sentinel_initial_identity', len(matches) == 1 and matches[0][2] is False
              and matches[0][3] == hashlib.sha256((sentinel_token+'\r').encode()).hexdigest())
        report['sentinel_before'] = matches[0]; flush()
        session.add_paragraph('Rollback original prefix 中文😀')
        session._execute_topology_mutation(['set targetRange to create range boundDoc start 0 end 8',
            f'make new bookmark at boundDoc with properties {{name:{apple_string(BOOKMARK)},text object:targetRange}}',
            'set nativeRows to {{"seeded"}}'])
        checkpoint = session.degradation_checkpoint()
        report['checkpoint'] = checkpoint
        saved = session._degradation_checkpoints[checkpoint]
        report['checkpoint_prefix_sha256'] = saved.state[0][7]; flush()
        # Native table/REF construction mirrors the already exercised recovery fixture.
        append = ['activate object boundWindow',
                  'set p to (end of content of text object of boundDoc) - 1',
                  'set r to create range boundDoc start p end p',
                  'set ownTable to make new table at boundDoc with properties {text object:r,number of rows:1,number of columns:1}',
                  'set p to start of content of text object of cell 1 of row 1 of ownTable',
                  'set r to create range boundDoc start p end p',
                  f'set content of r to {apple_string(TABLE_TEXT)}',
                  'set p to start of content of text object of cell 1 of row 1 of ownTable',
                  'set r to create range boundDoc start p end p',
                  f'create new field text range r field type field ref field text {apple_string(BOOKMARK)} preserve formatting true',
                  'set nativeRows to {{"appended",count fields of boundDoc,count tables of boundDoc}}']
        report['append_ack'] = session._execute_topology_mutation(append); flush()
        check('one_field_and_one_table_appended', report['append_ack'] == [['appended', 1, 1]])
        observed = session.snapshot_fields()
        check('field_cache_observed_before_rollback', len(observed) == 1 and hasattr(session, '_observed_field_topology'))
        tracking = tracking_state(session)
        report['owned_before_failure'] = independent_snapshot(output, 'owned-before-failure', session._bound_path, session._window_id, checkpoint)
        check('preflight_native_counts_and_prefix', report['owned_before_failure'][0][4:6] == [1, 1]
              and report['owned_before_failure'][0][7] == saved.state[0][7])
        scripts_before = script_state(session)
        with injected_compiler() as injections:
            try:
                session.rollback_degradation_checkpoint(checkpoint)
            except NativeWriterObjectError as error:
                report['public_rollback_error'] = {'type':type(error).__name__, 'code':error.code, 'message':str(error)}
            else:
                raise AssertionError('Injected rollback unexpectedly succeeded')
        report['injection'] = injections
        scripts_after = script_state(session)
        report['native_script_counts'] = {'before_rollback':len(scripts_before), 'after_rollback':len(scripts_after)}
        flush()
        check('single_injection_and_two_real_rollback_submissions', len(injections) == 1 and len(scripts_after) == len(scripts_before)+2)
        check('typed_public_error_and_quarantine', report['public_rollback_error']['code'] == 'LOCAL_MUTATION_ROLLBACK_FAILED'
              and session._quarantined and session._retain_evidence and session.lock.quarantine_path.is_file())
        check('observed_cache_invalidated_tracking_not_committed', not hasattr(session, '_observed_field_topology')
              and tracking == tracking_state(session))
        raw_logs = [path for path in session.staging_root.glob('*.log')
                    if 'WPSC_INJECTED_ROLLBACK_FAILURE (-2700)' in path.read_text()]
        report['ordinary_error_logs'] = [str(path) for path in raw_logs]; flush()
        check('raw_ordinary_native_error_retained', len(raw_logs) == 1)
        blocked = []
        for label, operation in (
                ('write', lambda: session.add_paragraph('FORBIDDEN POST-FAILURE WRITE')),
                ('context_close', lambda: session.__exit__(None, None, None))):
            try:
                operation()
            except NativeWordError as error:
                blocked.append({'operation':label, 'code':error.code})
            else:
                raise AssertionError(label+' was not blocked')
        report['blocked_operations'] = blocked
        report['native_script_counts']['after_blocked_write_and_close'] = len(script_state(session)); flush()
        check('later_write_and_context_close_emit_no_native_script', scripts_after == script_state(session)
              and all(row['code'] == 'NATIVE_WORD_QUARANTINED' for row in blocked) and not session._closed)
        # These diagnostic events are explicitly read-only and independent. Never
        # reset quarantine or rebind/retry a session operation to inspect the result.
        after = independent_snapshot(output, 'owned-after-quarantine', session._bound_path, session._window_id, checkpoint)
        report['owned_after_quarantine'] = after; flush()
        check('field_deleted_table_retained_prefix_unchanged', after[0][4:6] == [0, 1]
              and after[0][7] == saved.state[0][7] and len(after) == 2
              and after[1][0] == 'table' and after[1][5] is True)
        final_inventory = inventory(output, 'inventory-after-quarantine')
        report['inventory_after_quarantine'] = final_inventory; flush()
        check('unsaved_sentinel_exact_state_preserved', [row for row in final_inventory if row[0] == sentinel_name] == [report['sentinel_before']])
        owned_rows = [row for row in final_inventory if row[1] == session._bound_path]
        others = [row for row in final_inventory if row[0] != sentinel_name and row[1] != session._bound_path]
        check('owned_retained_and_unrelated_inventory_preserved', len(owned_rows) == 1 and others == report['inventory_before'])
        check('source_hashes_unchanged', all(sha(ROOT/name) == digest for name, digest in report['source_hashes'].items()))
        report['status'] = 'PASS_QUARANTINED_RETAINED'
    except BaseException as error:
        report['error'] = {'type':type(error).__name__, 'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and not session._quarantined:
            session._retain('Rollback failure fixture incomplete; retain exact owned document for manual cleanup')
    finally:
        report['source_digest_after'] = source_digest(ROOT)
        report['checks']['full_source_digest_unchanged'] = report['source_digest_before'] == report['source_digest_after']
        if not report['checks']['full_source_digest_unchanged']:
            report['status'] = 'FAIL_RETAINED'
            report['source_error'] = 'Shipped source digest changed during the native gate'
        if session:
            report['quarantined'] = session._quarantined
            report['closed'] = session._closed
            if session.staging_root and session.staging_root.exists():
                shutil.copytree(session.staging_root, output/'native-runtime', dirs_exist_ok=True)
            if session.lock:
                if session.lock.quarantine_path.is_file():
                    shutil.copyfile(session.lock.quarantine_path, output/'quarantine-marker.json')
                    report['quarantine_marker_sha256'] = sha(output/'quarantine-marker.json')
                session.lock.close()
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        parser.error('--execute is required; native document will remain quarantined for manual cleanup')
    report = run(args.output)
    print(json.dumps({'status':report['status'], 'output':str(Path(args.output).resolve())}))
    return 0 if report['status'] == 'PASS_QUARANTINED_RETAINED' else 1


if __name__ == '__main__':
    sys.exit(main())
