"""Explicit native Word session gate. Run only while Word's native job slot is free."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession
from skills.WPSComposer.scripts.msoffice.errors import NativeWordCapabilityError
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string


def run(output):
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {'passed': False, 'checks': {}, 'limitations': [
        'Attached non-rebinding save_copy and PDF export are unsupported.',
        'Floating object move/clone is unsupported.',
    ]}
    setup = MacWordSession()
    session = None
    sentinel = None
    try:
        setup._prepare()
        seed = setup.staging_root / 'seed.docx'
        token = 'WPSC-SENTINEL-' + uuid4().hex
        rows = setup._execute([
            'set seedDoc to make new document',
            'set content of text object of seedDoc to "Alpha" & return & "Beta" & return & "Gamma" & return',
            f'save as seedDoc file name {apple_string(str(seed))} file format format document default add to recent files false',
            'set seedDoc to document "seed.docx"',
            f'if posix full name of seedDoc is not {apple_string(str(seed))} then error "WPSC_SEED_BINDING"',
            'close seedDoc saving no',
            'set sentinelDoc to make new document',
            f'set content of text object of sentinelDoc to {apple_string(token)}',
            'set nativeRows to {{"binding", id of active window of sentinelDoc, posix full name of sentinelDoc as text, name of sentinelDoc as text}}',
        ], bind=False)
        sentinel = rows[0]
        setup.lock.close()
        source = output / 'source.docx'
        shutil.copy2(seed, source)
        original = hashlib.sha256(source.read_bytes()).hexdigest()
        report['source_sha256'] = original
        with MacWordSession.open_document(source) as session:
            before = session.inspect_document()
            report['before'] = before
            assert [p['text'] for p in before['paragraphs']][:3] == ['Alpha', 'Beta', 'Gamma']
            patch = session.apply_format_patch('paragraph:2', font={'name': 'Arial', 'size': 15, 'bold': True, 'color': '#123456'}, paragraph={'first_line_indent': 24, 'space_after': 8})
            assert not patch['rejected']
            after_format = session.inspect_document()
            assert after_format['paragraphs'][1]['font']['bold'] is True
            assert after_format['paragraphs'][1]['font']['size'] == 15
            assert after_format['paragraphs'][1]['paragraph']['first_line_indent'] == 24
            report['checks']['format_readback'] = True
            session.apply_structural_op({'op': 'insert', 'type': 'paragraph', 'position': {'before': 'paragraph:3'}, 'props': {'text': 'Inserted 中文'}})
            session.apply_structural_op({'op': 'clone', 'target': 'paragraph:1', 'to': 'end'})
            session.apply_structural_op({'op': 'insert', 'type': 'table', 'props': {'rows': 2, 'cols': 2, 'data': [['Header A', 'Header B'], ['甲', '乙']]}})
            session.apply_format_patch('table:1/cell:2,1', fill={'color': '#FF0000'}, vertical_alignment=1)
            session.apply_structural_op({'op': 'insert', 'type': 'textbox', 'props': {'text': 'Native Box', 'top': 350, 'width': 120, 'height': 40}})
            from PIL import Image
            image = output / 'image.png'
            Image.new('RGB', (24, 24), 'red').save(image)
            session.apply_structural_op({'op': 'insert', 'type': 'image', 'props': {'path': str(image)}})
            after = session.inspect_document()
            assert after['counts']['tables'] == 1
            assert after['counts']['shapes'] == 1
            assert after['counts']['inline_shapes'] == 1
            assert any(c['text'] == '甲' for c in after['tables'][0]['cells'])
            assert after['shapes'][0]['text'] == 'Native Box'
            assert sum(p['text'] == 'Alpha' for p in after['paragraphs']) == 2
            report['checks']['native_structural_objects'] = True
            report['after'] = after
            session.save_copy(output / 'edited.docx')
            session.export_pdf(output / 'edited.pdf')
            shutil.copytree(session.staging_root, output / 'runtime')
        session = None
        report['checks']['source_preserved'] = hashlib.sha256(source.read_bytes()).hexdigest() == original
        with MacWordSession.open_document(source) as editing:
            editing.apply_structural_op({'op': 'move', 'target': 'paragraph:3', 'to': 'start'})
            moved = editing.inspect_document()
            assert [p['text'] for p in moved['paragraphs']][:3] == ['Gamma', 'Alpha', 'Beta']
            editing.apply_structural_op({'op': 'remove', 'target': 'paragraph:2'})
            editing.apply_format_patch('paragraph:2', text='Beta revised')
            removed = editing.inspect_document()
            assert [p['text'] for p in removed['paragraphs']][:2] == ['Gamma', 'Beta revised']
            editing.save_copy(output / 'moved.docx')
            report['checks']['move_remove_and_text_boundary'] = True
        with MacWordSession.open_document(output / 'edited.docx', read_only=True) as reopened:
            snap = reopened.inspect_document()
            assert snap['paragraphs'][1]['font']['size'] == 15
            assert snap['counts']['tables'] == 1 and snap['counts']['shapes'] == 1
            report['checks']['reopen_persisted'] = True
        try:
            with MacWordSession.attach_active() as attached:
                assert attached._window_id == sentinel[1]
                selected = attached.inspect_selection()
                assert selected['id'] == 'selection'
                snap = attached.inspect_document()
                assert snap['paragraphs'][0]['text'] == token
                assert snap['saved'] is False
                report['checks']['attached_exact_unsaved_document'] = True
        except NativeWordCapabilityError:
            assert sentinel[1] is None
            report['limitations'].append('Native Word returned null window ID; active attachment fails before mutation.')
            report['checks']['unverifiable_attach_rejected'] = True
        import pdfplumber
        with pdfplumber.open(output / 'edited.pdf') as pdf:
            text = '\n'.join(p.extract_text() or '' for p in pdf.pages)
            assert 'Alpha' in text and 'Native Box' in text
        report['checks']['pdf_native_text'] = True
        report['passed'] = all(report['checks'].values())
    except BaseException as error:
        report['error'] = {'type': type(error).__name__, 'message': str(error)}
        (output / 'failure.txt').write_text(traceback.format_exc())
        if session is not None and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root, output / 'failed-runtime', dirs_exist_ok=True)
    finally:
        if setup.staging_root and setup.staging_root.exists():
            shutil.copytree(setup.staging_root, output / 'setup-runtime', dirs_exist_ok=True)
        if sentinel is not None:
            try:
                setup.lock.acquire(__import__('time').monotonic() + 10)
                setup._execute([
                    f'set boundDoc to document {apple_string(sentinel[3])}',
                    f'if (posix full name of boundDoc as text) is not {apple_string(sentinel[2])} then error "WPSC_SENTINEL_PATH_CHANGED"',
                    f'if (content of text object of boundDoc as text) is not {apple_string(token)} & return then error "WPSC_SENTINEL_CHANGED"',
                    'if saved of boundDoc then error "WPSC_SENTINEL_SAVED"',
                    'close boundDoc saving no', 'set nativeRows to {{"ok"}}',
                ], bind=False)
                report['checks']['sentinel_preserved_and_closed'] = True
            except BaseException as error:
                report['passed'] = False
                report['cleanup_error'] = str(error)
        setup.lock.close() if setup.lock else None
        if setup.staging_root and setup.staging_root.exists() and not setup._quarantined and report.get('checks', {}).get('sentinel_preserved_and_closed'):
            shutil.rmtree(setup.staging_root)
        (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2))
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', type=Path, required=True)
    args = parser.parse_args()
    result = run(args.output_dir)
    print(json.dumps({'passed': result['passed'], 'checks': result['checks'], 'error': result.get('error')}, ensure_ascii=False))
    raise SystemExit(0 if result['passed'] else 1)
