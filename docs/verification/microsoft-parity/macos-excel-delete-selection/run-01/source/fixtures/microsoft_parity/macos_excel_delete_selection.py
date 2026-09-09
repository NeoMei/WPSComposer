"""Opt-in native acceptance for logical selection after empty-sheet deletion."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import traceback
from uuid import uuid4

from skills.WPSComposer import open_document
from skills.WPSComposer.scripts.msoffice.macos_excel_session import _quote
from fixtures.microsoft_parity.macos_excel_integration_round2 import (
    ROOT, _inventory, _osascript, _sha,
)


def run(source, output):
    source = Path(source).resolve()
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    files = list((ROOT / 'skills/WPSComposer').rglob('*.py')) + [
        Path(__file__).resolve(),
        ROOT / 'fixtures/microsoft_parity/macos_excel_integration_round2.py',
    ]
    hashes = {str(p.relative_to(ROOT)): _sha(p) for p in files}
    report = {'passed': False, 'source_sha256': _sha(source),
              'source_hashes': hashes, 'checks': {}, 'jobs': []}
    for path in files[-2:] + [ROOT / 'skills/WPSComposer/scripts/msoffice/macos_excel_session.py']:
        dest = output / 'source' / path.relative_to(ROOT)
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, dest)
    before = None
    sentinel = None
    session = None
    marker = 'WPSC-DELETE-SENTINEL-' + uuid4().hex
    try:
        before = _inventory(output, 'inventory-before')
        sentinel = _osascript(output, 'create-sentinel', f'''set sentinelBook to make new workbook
set value of range "A1" of worksheet 1 of sentinelBook to {_quote(marker)}
return name of sentinelBook''')
        sentinel_inventory = _inventory(output, 'inventory-with-sentinel')
        assert len(sentinel_inventory) == len(before) + 1
        for case in ('before-selected', 'selected', 'after-selected'):
            session = open_document(source, engine='msoffice')
            report['jobs'].append(str(session._job))
            initial = session.inspect_document(include_cells=False)
            base_count = initial['sheet_count']
            assert base_count >= 1
            a = session.add_sheet('DeleteProbeA')
            b = session.add_sheet('DeleteProbeB')
            assert a['path'] == f'sheet:{base_count + 1}'
            assert b['path'] == f'sheet:{base_count + 2}'
            if case == 'before-selected':
                target, expected = a['path'], 'DeleteProbeB'
            elif case == 'selected':
                target, expected = b['path'], 'DeleteProbeA'
            else:
                session.select_sheet(1)
                target, expected = b['path'], initial['sheets'][0]['name']
            session.apply_structural_op({'op': 'remove', 'target': target})
            value = 'AFTER DELETE ' + case
            session.write_cell(1, 1, value)
            snap = session.inspect_document(max_cells=100)
            matches = [s['name'] for s in snap['sheets'] if any(
                c['address'] == '$A$1' and c['value'] == value for c in s['cells'])]
            assert matches == [expected], matches
            assert snap['sheet_count'] == base_count + 1
            path = output / (case + '.xlsx')
            session.save(path)
            session.close(save_changes=False)
            session = None
            session = open_document(path, read_only=True, engine='msoffice')
            report['jobs'].append(str(session._job))
            reopened = session.inspect_document(max_cells=100)
            assert [s['name'] for s in reopened['sheets'] if any(
                c['address'] == '$A$1' and c['value'] == value for c in s['cells'])] == [expected]
            assert reopened['sheet_count'] == base_count + 1
            (output / (case + '-reopened.json')).write_text(json.dumps(reopened, indent=2) + '\n')
            session.close(save_changes=False)
            session = None
            assert _sha(source) == report['source_sha256']
            report['checks'][case + '-write-save-reopen'] = True
        assert _inventory(output, 'inventory-after-owned-close') == sentinel_inventory
        report['checks']['owned-workbooks-closed'] = True
        report['checks']['source-bytes-preserved'] = _sha(source) == report['source_sha256']
        report['checks']['source-hashes-unchanged'] = all(_sha(ROOT / p) == h for p, h in hashes.items())
        report['passed'] = all(report['checks'].values())
    except BaseException as error:
        report['error'] = {'type': type(error).__name__, 'message': str(error)}
        (output / 'failure.txt').write_text(traceback.format_exc())
    finally:
        if session is not None and not session._closed and not session._failed:
            try:
                session.close(save_changes=False)
            except BaseException as error:
                report['cleanup_error'] = str(error)
                report['passed'] = False
        if sentinel is not None:
            try:
                state = _osascript(output, 'verify-sentinel', f'''set sentinelBook to workbook {_quote(sentinel)}
return ((saved of sentinelBook) as text) & (character id 31) & ((value of range "A1" of worksheet 1 of sentinelBook) as text)''').split(chr(31))
                assert state == ['false', marker]
                report['checks']['unsaved-sentinel-preserved'] = True
                _osascript(output, 'close-exact-sentinel', f'''close workbook {_quote(sentinel)} saving no
return "closed exact sentinel"''')
            except BaseException as error:
                report['sentinel_cleanup_error'] = str(error)
                report['passed'] = False
        if before is not None:
            try:
                report['inventory_restored'] = _inventory(output, 'inventory-final') == before
                report['passed'] = report['passed'] and report['inventory_restored']
            except BaseException as error:
                report['inventory_error'] = str(error)
                report['passed'] = False
        for index, job in enumerate(report['jobs']):
            shutil.copytree(job, output / 'runtime' / str(index), dirs_exist_ok=True)
        (output / 'report.json').write_text(json.dumps(report, indent=2) + '\n')
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--execute-native', action='store_true')
    args = parser.parse_args()
    if not args.execute_native:
        parser.error('--execute-native is required')
    result = run(args.source, args.output)
    print(json.dumps({k: v for k, v in result.items() if k not in ('source_hashes', 'jobs')}, indent=2))
    raise SystemExit(0 if result['passed'] else 1)
