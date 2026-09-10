"""One native DOCX insert-file control. Source preparation only, opt-in execution.

Changes only the donor format relative to the failed Flat OPC import. A native
Word-saved DOCX is copied to the recipient's private runtime; command, range,
file-path convention and flags are the reviewed import fixture's exact helper.
No fallback, retry, public implementation or Flat OPC acceptance follows.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession, apple_string, sha, inventory, retain_sources,
    sentinel_preimage, close_after_owned, import_commands,
    replacement_snapshot, replacement_expectation, seed, SOURCES,
)

DONOR_TEXT = 'DOCX CONTROL 中文😀'


def run(output):
    report = {'status': 'FAIL', 'scope': 'One Word-saved DOCX insert-file control only',
              'source_hashes': retain_sources(output, [Path(__file__), *SOURCES]),
              'checks': {}, 'raw_rows': {}}
    def flush():
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    def check(key, value):
        report['checks'][key] = bool(value)
        flush()
        if not value:
            raise AssertionError('DOCX control failed: '+key)
    def execute(owner, label, commands):
        report['current_phase'] = label
        flush()
        rows = owner._execute(commands)
        report['raw_rows'][label] = rows
        flush()
        return rows
    def owned_stage(label, action):
        directory = output/label
        directory.mkdir()
        detail = {'checks': {}, 'inventory_before': inventory(directory, 'before')}
        report[label] = detail
        owner = None
        sentinel_name = token = None
        try:
            with MacWordSession.new_document(visible=False) as owner:
                owner._retain_evidence = True
                token = 'DOCX CONTROL SENTINEL '+uuid4().hex+' 中文😀'
                rows = execute(owner, label+'_sentinel', ['set sentinelDoc to make new document',
                               f'set content of text object of sentinelDoc to {apple_string(token)}',
                               'set nativeRows to {{name of sentinelDoc as text}}'])
                if len(rows) != 1 or len(rows[0]) != 1 or not isinstance(rows[0][0], str):
                    raise ValueError('Sentinel acknowledgement invalid')
                sentinel_name = rows[0][0]
                detail['sentinel_name'] = sentinel_name
                detail['sentinel_before'] = sentinel_preimage(inventory(directory, 'sentinel-before'), sentinel_name, token)
                report['word_version'] = execute(owner, label+'_version', ['set nativeRows to {{version as text}}'])
                action(owner)
            close_after_owned(owner, directory, detail, sentinel_name, token)
            sentinel_name = None
        finally:
            if owner and sentinel_name and not detail.get('sentinel_cleanup_attempted'):
                try:
                    close_after_owned(owner, directory, detail, sentinel_name, token)
                    sentinel_name = None
                except BaseException:
                    detail['cleanup_failure'] = traceback.format_exc()
            if owner and owner.staging_root and owner.staging_root.exists():
                shutil.copytree(owner.staging_root, directory/'native-runtime', dirs_exist_ok=True)
            detail['remaining_sentinel'] = sentinel_name
            flush()
        check(label+'_cleanup_verified', sentinel_name is None and not detail.get('cleanup_failure') and all(detail['checks'].values()))
    try:
        report['dictionary_sha256'] = sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before'] = inventory(output, 'initial')
        donor_path = output/'native-donor.docx'
        def donor(owner):
            rows = execute(owner, 'donor_seed', [f'set content of text object of boundDoc to {apple_string(DONOR_TEXT)}',
                           'set nativeRows to {{content of text object of boundDoc as text}}'])
            check('donor_native_text', rows == [[DONOR_TEXT+'\r']])
            report['donor_text'] = rows[0][0]
            owner.save_docx(donor_path)
            report['donor_sha256'] = sha(donor_path)
        owned_stage('donor', donor)
        check('inventory_after_donor', inventory(output, 'after-donor') == report['inventory_before'])
        def recipient(owner):
            check('recipient_seed', execute(owner, 'recipient_seed', seed()) == [['stage', True]])
            owner.save_docx(output/'recipient-preimage.docx')
            report['recipient_preimage_sha256'] = sha(output/'recipient-preimage.docx')
            native_donor = owner.staging_root/'heading-input.docx'
            shutil.copyfile(donor_path, native_donor)
            check('donor_input_identical', sha(native_donor) == report['donor_sha256'])
            before = execute(owner, 'recipient_preimage', replacement_snapshot())
            check('preimage_shape', len(before) == 1 and len(before[0]) == 4 and before[0][0] == 'preimage')
            _, text, start, end = before[0]
            expected = replacement_expectation(text, start, end, report['donor_text'])
            report['expected_text'] = expected
            rows = execute(owner, 'single_native_docx_import', import_commands(native_donor))
            check('exact_docx_replacement', rows == [['import', start, end, expected]])
            owner.save_docx(output/'recipient-after.docx')
            check('native_input_preserved', sha(native_donor) == report['donor_sha256'])
        owned_stage('recipient', recipient)
        check('donor_file_preserved', sha(donor_path) == report['donor_sha256'])
        check('recipient_preimage_preserved', sha(output/'recipient-preimage.docx') == report['recipient_preimage_sha256'])
        report['inventory_final'] = inventory(output, 'final')
        check('final_inventory_preserved', report['inventory_final'] == report['inventory_before'])
        check('sources_unchanged', all(sha(ROOT/path) == digest for path, digest in report['source_hashes'].items()))
        report['status'] = 'PASS'
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix == '.docx'}
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('DOCX control requires explicit --execute', file=sys.stderr)
        return 2
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    return 0 if run(output)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
