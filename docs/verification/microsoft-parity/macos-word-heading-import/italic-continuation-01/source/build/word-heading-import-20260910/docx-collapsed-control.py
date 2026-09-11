"""One collapsed native DOCX import, then guarded old-target removal.

Fixture only. Reuse the immutable Word-saved donor from docx-control-01.
No Flat OPC opening, retry, text rollback, product method or capability enable.
"""
from __future__ import annotations
import argparse
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
from fixtures.microsoft_parity.macos_word_heading_import import (
    MacWordSession, apple_string, sha, inventory, retain_sources,
    sentinel_preimage, close_after_owned, replacement_snapshot,
    replacement_expectation, seed, SOURCES, existing_styles_preserved,
    package_styles, canonical, W,
)
DONOR = Path(__file__).parent/'docx-control-01/native-donor.docx'
DONOR_SHA = 'd7298013c60a9b0166cd6c1efce2d8a829e590fd92633443f5e9787274cd1cb8'
DONOR_TEXT = 'DOCX CONTROL 中文😀\r'


def units(text):
    return len(text.encode('utf-16-le'))//2


def expected_stages(text, start, end, donor):
    final = replacement_expectation(text, start, end, donor)
    encoded = text.encode('utf-16-le')
    inserted = encoded[:start*2].decode('utf-16-le') + donor + encoded[start*2:].decode('utf-16-le')
    return inserted, final, start+units(donor), end+units(donor)


def exact_ack(rows, label, start, end, text):
    return (isinstance(rows, list) and len(rows) == 1 and isinstance(rows[0], list)
            and len(rows[0]) == 4 and type(rows[0][1]) is int and type(rows[0][2]) is int
            and rows == [[label, start, end, text]])


def full_guard(text):
    return [f'if not (((current application\'s NSString\'s stringWithString:(content of text object of boundDoc as text))\'s isEqualToString:{apple_string(text)}) as boolean) then error "WPSC_COLLAPSED_FULL_PREIMAGE"']


def collapsed_commands(path, start, preimage):
    if type(start) is not int or start < 0:
        raise ValueError('Invalid collapsed coordinate')
    return full_guard(preimage) + [
        f'set importRange to create range boundDoc start {start} end {start}',
        'set importStart to start of content of importRange',
        'set importEnd to end of content of importRange',
        f'if importStart is not {start} or importEnd is not {start} then error "WPSC_COLLAPSED_RANGE"',
        f'insert file at importRange file name {apple_string(str(path))} confirm conversions false link false',
        'set nativeRows to {{"insert",importStart,importEnd,content of text object of boundDoc as text}}',
    ]


def clear_commands(preimage, start, end):
    replacement_expectation(preimage, start, end, '')
    return full_guard(preimage) + [
        f'set clearRange to create range boundDoc start {start} end {end}',
        'if content of clearRange is not "REPLACE" & return then error "WPSC_COLLAPSED_TARGET_CHANGED"',
        'set clearStart to start of content of clearRange',
        'set clearEnd to end of content of clearRange',
        f'if clearStart is not {start} or clearEnd is not {end} then error "WPSC_COLLAPSED_TARGET_BOUNDS"',
        'set content of clearRange to ""',
        'set nativeRows to {{"clear",clearStart,clearEnd,content of text object of boundDoc as text}}',
    ]


def unrelated_inventory_preserved(before, after, owned_path):
    def split(rows):
        if not isinstance(rows, list) or any(not isinstance(r, list) or len(r) != 4 for r in rows):
            return None
        own = [r for r in rows if r[1] == str(owned_path)]
        return (own, [r for r in rows if r[1] != str(owned_path)]) if len(own) == 1 else None
    old, new = split(before), split(after)
    return bool(old and new and old[0][0][:2] == new[0][0][:2] and old[1] == new[1])


def styles_allowed(before, after, donor):
    if not existing_styles_preserved(before, after):
        return False
    def index(xml):
        return {s.get(W+'styleId'): canonical(s) for s in ET.fromstring(xml).findall(W+'style')}
    old, new, source = index(before), index(after), index(donor)
    # Only exact definitions already carried by the native donor may be added.
    return all(identifier in old or source.get(identifier) == value for identifier, value in new.items())


def run(output):
    report = {'status': 'FAIL', 'scope': 'Collapsed DOCX import/delete feasibility only',
              'source_hashes': retain_sources(output, [Path(__file__), Path(__file__).with_name('test_docx_collapsed_control.py'), *SOURCES]),
              'checks': {}, 'raw_rows': {}}
    owner = None
    sentinel_name = token = None
    def flush():
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    def check(label, value):
        report['checks'][label] = bool(value)
        flush()
        if not value:
            raise AssertionError('Collapsed control failed: '+label)
    def execute(label, commands):
        report['current_phase'] = label
        flush()
        rows = owner._execute(commands)
        report['raw_rows'][label] = rows
        flush()
        return rows
    try:
        check('known_word_saved_donor', sha(DONOR) == DONOR_SHA)
        donor_styles = package_styles(DONOR)
        with ZipFile(DONOR) as package:
            body = ET.fromstring(package.read('word/document.xml')).find(W+'body')
            actual_donor = ''.join(''.join(t.text or '' for t in p.iter(W+'t'))+'\r' for p in body.findall(W+'p'))
        check('known_donor_text', actual_donor == DONOR_TEXT)
        report['donor_sha256'] = DONOR_SHA
        report['dictionary_sha256'] = sha('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef')
        report['inventory_before'] = inventory(output, 'before')
        with MacWordSession.new_document(visible=False) as owner:
            owner._retain_evidence = True
            token = 'COLLAPSED SENTINEL '+uuid4().hex+' 中文😀'
            rows = execute('sentinel', ['set sentinelDoc to make new document',
                           f'set content of text object of sentinelDoc to {apple_string(token)}',
                           'set nativeRows to {{name of sentinelDoc as text}}'])
            check('sentinel_ack', len(rows) == 1 and len(rows[0]) == 1 and isinstance(rows[0][0], str))
            sentinel_name = rows[0][0]
            report['sentinel_name'] = sentinel_name
            report['sentinel_before'] = sentinel_preimage(inventory(output, 'sentinel-before'), sentinel_name, token)
            report['word_version'] = execute('version', ['set nativeRows to {{version as text}}'])
            check('seed_ack', execute('seed', seed()) == [['stage', True]])
            owner.save_docx(output/'recipient-preimage.docx')
            report['preimage_sha256'] = sha(output/'recipient-preimage.docx')
            before_styles = package_styles(output/'recipient-preimage.docx')
            native_donor = owner.staging_root/'heading-input.docx'
            shutil.copyfile(DONOR, native_donor)
            check('copied_donor_identical', sha(native_donor) == DONOR_SHA)
            before = execute('native_preimage', replacement_snapshot())
            check('preimage_shape', len(before) == 1 and len(before[0]) == 4 and before[0][0] == 'preimage')
            _, text, start, end = before[0]
            inserted, final, shifted_start, shifted_end = expected_stages(text, start, end, DONOR_TEXT)
            report['expected'] = {'inserted': inserted, 'final': final, 'shifted_start': shifted_start, 'shifted_end': shifted_end}
            before_inventory = inventory(output, 'before-insert')
            report['inventory_before_insert'] = before_inventory
            rows = execute('single_collapsed_import', collapsed_commands(native_donor, start, text))
            check('exact_collapsed_insert', exact_ack(rows, 'insert', start, start, inserted))
            after_inventory = inventory(output, 'after-insert')
            report['inventory_after_insert'] = after_inventory
            check('insert_inventory_preserved', unrelated_inventory_preserved(before_inventory, after_inventory, owner._bound_path))
            owner.save_docx(output/'after-insert.docx')
            inserted_styles = package_styles(output/'after-insert.docx')
            check('insert_styles_preserved', styles_allowed(before_styles, inserted_styles, donor_styles))
            # Deletion is reachable only after text, inventory and style ACKs.
            rows = execute('clear_shifted_original', clear_commands(inserted, shifted_start, shifted_end))
            check('exact_clear', exact_ack(rows, 'clear', shifted_start, shifted_end, final))
            after_inventory = inventory(output, 'after-clear')
            report['inventory_after_clear'] = after_inventory
            check('clear_inventory_preserved', unrelated_inventory_preserved(before_inventory, after_inventory, owner._bound_path))
            owner.save_docx(output/'recipient-after.docx')
            final_styles = package_styles(output/'recipient-after.docx')
            check('clear_styles_preserved', styles_allowed(before_styles, final_styles, donor_styles) and existing_styles_preserved(inserted_styles, final_styles))
            check('native_donor_preserved', sha(native_donor) == DONOR_SHA)
        close_after_owned(owner, output, report, sentinel_name, token)
        sentinel_name = None
        check('original_donor_preserved', sha(DONOR) == DONOR_SHA)
        check('source_snapshot_preserved', sha(output/'recipient-preimage.docx') == report['preimage_sha256'])
        report['inventory_final'] = inventory(output, 'final')
        check('final_inventory_preserved', report['inventory_final'] == report['inventory_before'])
        check('sources_unchanged', all(sha(ROOT/path) == digest for path, digest in report['source_hashes'].items()))
        report['status'] = 'PASS'
    except BaseException as exc:
        report['error'] = {'type': type(exc).__name__, 'message': str(exc)}
        (output/'failure.txt').write_text(traceback.format_exc())
    finally:
        if owner and sentinel_name and not report.get('sentinel_cleanup_attempted'):
            try:
                close_after_owned(owner, output, report, sentinel_name, token)
                sentinel_name = None
            except BaseException:
                report['cleanup_failure'] = traceback.format_exc()
        if owner and owner.staging_root and owner.staging_root.exists():
            shutil.copytree(owner.staging_root, output/'native-runtime', dirs_exist_ok=True)
        if sentinel_name or report.get('cleanup_failure'):
            report['status'] = 'FAIL'
        report['remaining_sentinel'] = sentinel_name
        report['artifact_hashes'] = {p.name: sha(p) for p in output.iterdir() if p.suffix == '.docx'}
        flush()
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args(argv)
    if not args.execute:
        print('Collapsed control requires explicit --execute', file=sys.stderr)
        return 2
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    return 0 if run(output)['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
