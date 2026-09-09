"""Guarded pagination acceptance. Run only while holding the root's Word lease."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from types import SimpleNamespace
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer import create_document, open_document
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from fixtures.microsoft_parity.macos_word_fields import retain_sources
from fixtures.microsoft_parity.macos_word_recovery import inventory, close_sentinel


def _digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _state(session):
    # Hash inside the native script, retaining no document content in logs.
    return session._execute([
        'set paginationText to content of text object of boundDoc as text',
        'set paginationHash to do shell script ("/usr/bin/printf %s " & quoted form of paginationText & " | /usr/bin/shasum -a 256")',
        'set nativeRows to {{"state",paginationHash,end of content of text object of boundDoc,'
        'count paragraphs of boundDoc,count tables of boundDoc,count bookmarks of boundDoc,'
        'count fields of boundDoc,saved of boundDoc}}',
    ])


def _tracked(content_end):
    return [
        {'nodeId':'whole😀','op':'writer.add_heading','role':'正文',
         'range':SimpleNamespace(Start=0,End=content_end)},
        {'nodeId':'emoji😀','range':SimpleNamespace(Start=0,End=2)},
        {'nodeId':'end','op':'writer.add_equation',
         'range':SimpleNamespace(Start=content_end-1,End=content_end-1)},
        {'nodeId':'past-end','op':'writer.add_semantic_table',
         'range':SimpleNamespace(Start=content_end+5,End=content_end+9)},
        {'nodeId':'/private/client/secret.docx',
         'range':SimpleNamespace(Start=0,End=0)},
        {'nodeId':'whole😀'}, {'nodeId':None},
    ]


def _snapshots(session, content_end):
    return {
        'map': session.pagination_map_for_ranges(_tracked(content_end)),
        'bookmark': session.pagination_fragment_for_bookmark('first😀','WPSC_Pagination'),
        'ordinal': session.pagination_fragment_for_bookmark('first😀',1),
    }


def _validate(snapshots, content_end):
    result = snapshots['map']
    assert result['version'] == 'M5-v1'
    nodes = result['nodes']
    assert [n['nodeId'] for n in nodes] == ['whole😀','emoji😀','end','past-end','<redacted>']
    assert nodes[0]['pageStart'] == 1 and nodes[0]['pageEnd'] == 3
    assert [f['page'] for f in nodes[0]['fragments']] == [1,2,3]
    assert nodes[0]['sections'] == ['正文']
    assert nodes[1]['range'] == '0:2' and nodes[1]['fragments'] == [{'page':1}]
    assert nodes[2]['pageStart'] == nodes[2]['pageEnd'] == 3
    assert nodes[3]['pageStart'] == nodes[3]['pageEnd'] == 3
    assert nodes[3]['range'] == f'{content_end+5}:{content_end+9}'
    assert nodes[2]['fragments'] == nodes[3]['fragments']
    # Coordinates may be unavailable (-1) in a hidden native window. That is
    # frozen semantic fallback, not a reason to invent geometry.
    for node in nodes:
        for fragment in node['fragments']:
            if 'bounds' in fragment:
                x0,y0,x1,y1 = fragment['bounds']
                assert x0 >= 0 and y0 >= 0 and x1 > x0 and y1 > y0
    bookmark = snapshots['bookmark']
    assert bookmark == snapshots['ordinal']
    assert bookmark['pageStart'] == bookmark['pageEnd'] == 1
    assert bookmark['range'].startswith('0:') and int(bookmark['range'].split(':')[1]) > 2
    if 'bounds' in bookmark['fragments'][0]:
        x0,y0,x1,y1 = bookmark['fragments'][0]['bounds']
        assert x1-x0 == 1 and y1-y0 == 12
    serialized = json.dumps(snapshots,ensure_ascii=False)
    assert 'First page body' not in serialized and '/private/client' not in serialized


def run(output):
    output = Path(output).resolve()
    output.mkdir(parents=True,exist_ok=False)
    sources = [Path(__file__), ROOT/'skills/WPSComposer/__init__.py',
               ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_session.py',
               ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_pagination.py',
               ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_degradation.py',
               ROOT/'skills/WPSComposer/scripts/longform/privacy.py',
               ROOT/'skills/WPSComposer/scripts/writer.py',
               ROOT/'fixtures/microsoft_parity/macos_word_fields.py',
               ROOT/'fixtures/microsoft_parity/macos_word_recovery.py',
               ROOT/'skills/WPSComposer/scripts/msoffice/macos_word_recovery.py']
    report = {'passed':False,'source_hashes':retain_sources(output,sources),'checks':{}}
    session = None
    sentinel = None
    try:
        starting = inventory(output,'starting-inventory')
        report['starting_inventory'] = starting
        with create_document('writer',engine='msoffice',visible=False) as session:
            token = 'Pagination sentinel 中文😀 '+uuid4().hex
            sentinel = session._execute([
                'set sentinelDoc to make new document',
                f'set content of text object of sentinelDoc to {apple_string(token)}',
                'set nativeRows to {{name of sentinelDoc as text}}',
            ])[0][0]
            before_owned = inventory(output,'before-owned-write')
            unrelated = [r for r in before_owned if r[1] != session._bound_path]
            report['inventory_before'] = unrelated
            sentinel_before = next(r for r in unrelated if r[0] == sentinel)
            assert sentinel_before[2] is False
            report['sentinel_before'] = sentinel_before
            assert sorted(r for r in unrelated if r[0] != sentinel) == starting
            try:
                session.add_heading_level('😀 分页 First page',1)
                session.add_paragraph('First page body 中文 é 𠀀.')
                session.add_page_break()
                session.add_paragraph('Second page body 中文😀.')
                session.add_page_break()
                session.add_paragraph('Third page body 中文😀.')
                session._execute([
                    'set paginationAnchor to create range boundDoc start 0 end 2',
                    'make new bookmark at boundDoc with properties {name:"WPSC_Pagination",text object:paginationAnchor}',
                    'set nativeRows to {{"ok"}}',
                ])
                before = _state(session)
                content_end = before[0][2]
                snapshots = _snapshots(session,content_end)
                _validate(snapshots,content_end)
                assert _state(session) == before
                assert _snapshots(session,content_end) == snapshots
                report['snapshot'] = snapshots
                report['native_state_before'] = before
                report['native_state_after'] = _state(session)
                report['checks']['native_read_only_pagination_no_text_or_saved_state_mutation'] = True
                report['checks']['native_unicode_multpage_dedup_clamping_privacy'] = True
                try:
                    session.pagination_fragment_for_bookmark('private','WPSC_Missing')
                except Exception as error:
                    assert getattr(error,'code',None) == 'PAGINATION_SNAPSHOT_FAILED'
                    assert str(error) == 'pagination snapshot failed'
                else:
                    raise AssertionError('Missing native bookmark was accepted')
                assert _state(session) == before
                report['checks']['missing_native_bookmark_typed_without_mutation'] = True
                session.save_docx(output/'pagination.docx')
                session.export_pdf(output/'pagination.pdf')
                report['checks']['public_native_save_and_pdf'] = True
                after_save = inventory(output,'after-save')
                assert [r for r in after_save if r[1] != session._bound_path] == unrelated
                report['checks']['unrelated_unsaved_documents_preserved'] = True
            finally:
                if session.staging_root.exists():
                    shutil.copytree(session.staging_root,output/'native-runtime',dirs_exist_ok=True)
        after_owned = inventory(output,'after-owned-close')
        report['inventory_after_owned_close'] = after_owned
        assert next(r for r in after_owned if r[0] == sentinel) == sentinel_before
        assert [r for r in after_owned if r[0] != sentinel] == starting
        report['checks']['owned_context_closed_with_unsaved_sentinel_preserved'] = True
        source = output/'pagination.docx'
        digest = _digest(source)
        with open_document(source,engine='msoffice',read_only=True,visible=False) as session:
            assert session._read_only
            before = _state(session)
            report['reopen_snapshot'] = _snapshots(session,content_end)
            assert report['reopen_snapshot'] == snapshots
            assert _state(session) == before
            report['checks']['read_only_native_reopen_same_pagination'] = True
            shutil.copytree(session.staging_root,output/'reopen-runtime')
        assert _digest(source) == digest
        report['checks']['source_bytes_preserved_after_read_only_reopen'] = True
        after_reopen = inventory(output,'after-reopen-close')
        report['inventory_after_reopen_close'] = after_reopen
        assert next(r for r in after_reopen if r[0] == sentinel) == sentinel_before
        assert [r for r in after_reopen if r[0] != sentinel] == starting
        report['checks']['reopen_context_closed_with_unsaved_sentinel_preserved'] = True
        report['sentinel_close'] = close_sentinel(output,sentinel,token)
        sentinel = None
        report['final_inventory'] = inventory(output,'after-sentinel-close')
        assert report['final_inventory'] == starting
        report['checks']['independent_sentinel_close_restores_starting_inventory'] = True
        report['source_sha256_before_reopen'] = digest
        report['source_sha256_after_reopen'] = _digest(source)
        import pdfplumber
        with pdfplumber.open(output/'pagination.pdf') as pdf:
            assert len(pdf.pages) == 3
            for page, text in zip(pdf.pages, ('First page body','Second page body','Third page body')):
                assert text in (page.extract_text() or '')
            report['pdf_page_count'] = len(pdf.pages)
        report['checks']['independent_pdf_three_page_text_placement'] = True
        report['artifact_hashes'] = {name:_digest(output/name) for name in ('pagination.docx','pagination.pdf')}
        report['passed'] = all(report['checks'].values())
    except BaseException as error:
        report['error'] = {'type':type(error).__name__,'message':str(error)}
        (output/'failure.txt').write_text(traceback.format_exc())
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root,output/'failed-runtime',dirs_exist_ok=True)
        if sentinel:
            report['retained_sentinel_name'] = sentinel
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute',action='store_true')
    parser.add_argument('--output-dir',required=True)
    args = parser.parse_args()
    if not args.execute:
        parser.error('--execute required; root must hold the Word native lease')
    result = run(args.output_dir)
    print(json.dumps({'passed':result['passed'],'error':result.get('error')}))
    sys.exit(0 if result['passed'] else 1)
