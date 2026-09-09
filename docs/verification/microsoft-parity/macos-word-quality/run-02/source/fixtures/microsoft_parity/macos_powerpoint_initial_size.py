"""Bounded public macOS PowerPoint initial-size acceptance.

Runs public create_document with a tested non-default initial size and public
generate with the unchanged default size. Inventory evidence hashes unrelated
text and requires an exact unsaved sentinel. No GUI, clipboard, Quit or static
OOXML mutation is used.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import time
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from fixtures.microsoft_parity.macos_powerpoint_size_probe import _inventory


ROOT = Path(__file__).resolve().parents[2]


def _digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _pptx_size(path):
    with ZipFile(path) as package:
        root = ET.fromstring(package.read('ppt/presentation.xml'))
    size = root.find('{http://schemas.openxmlformats.org/presentationml/2006/main}sldSz')
    return [int(size.attrib['cx']) / 12700, int(size.attrib['cy']) / 12700]


def _pdf(path):
    import fitz
    with fitz.open(path) as document:
        return {
            'pages': [[page.rect.width, page.rect.height] for page in document],
            'text': '\n'.join(page.get_text() for page in document),
        }


def _copy_job(session, destination):
    job = Path(session._job)
    if not job.is_dir():
        raise RuntimeError('Native PowerPoint job evidence is unavailable')
    shutil.copytree(job, destination)
    scripts = sorted(destination.glob('*.applescript'))
    if not scripts or 'close ownedDoc saving no' not in scripts[-1].read_text():
        raise RuntimeError('Exact owned close evidence is unavailable')


def _check_sentinel(rows, name):
    matches = [row for row in rows if row[0] == name]
    if len(matches) != 1 or matches[0][2] is not False or not matches[0][4]['has_text']:
        raise ValueError('Sentinel must be one exact unsaved presentation with text')


def _create_variant(output, width, height, *, create_document, open_document, deadline):
    label = f'{width} x {height}'
    session = create_document('slide', engine='msoffice', visible=False)
    session.timeout = min(90, deadline - time.monotonic())
    with session:
        session.set_slide_size(width, height)
        _, index = session.add_blank_slide()
        target = session.add_textbox(index, 'Public create ' + label, 30, 30, 260, 40)
        session.apply_format_patch(target, text_frame={'auto_size': 0},
                                   geometry={'left': 30, 'top': 30, 'width': 260, 'height': 40})
        session.add_shape(index, 1, 40, 110, 180, 60,
                          fill_color='#2463A6', text='Native rectangle')
        created = session.inspect_document()
        rejection = None
        try:
            session.set_slide_size(height, width)
        except ValueError as exc:
            rejection = str(exc)
        after_rejection = session.inspect_document()
        session.save(output / 'created.pptx')
        session.export_pdf(output / 'created.pdf')
    _copy_job(session, output / 'native-job')
    source_hash_before_reopen = _digest(output / 'created.pptx')
    with open_document(output / 'created.pptx', kind='slide',
                       engine='msoffice', read_only=True) as reopened:
        reopened.timeout = min(90, deadline - time.monotonic())
        snapshot = reopened.inspect_document()
    _copy_job(reopened, output / 'reopen-job')
    pdf = _pdf(output / 'created.pdf')
    textbox = next(shape for shape in snapshot['slides'][0]['shapes']
                   if shape['text'] == 'Public create ' + label)
    checks = {
        'size': _pptx_size(output / 'created.pptx') == [width, height],
        'reopen_width': abs(snapshot['page_setup']['slide_width'] - width) < .1,
        'geometry': all(abs(textbox['geometry'][key] - value) < .1
                        for key, value in {'left': 30, 'top': 30,
                                           'width': 260, 'height': 40}.items()),
        'pdf': (len(pdf['pages']) == 1
                and all(abs(a - b) < .1 for a, b in zip(pdf['pages'][0], [width, height]))
                and 'Public create ' + label in pdf['text']
                and 'Native rectangle' in pdf['text']),
        'existing_content_rejected_without_change': (
            rejection == 'Arbitrary initial slide size requires an empty presentation'
            and created == after_rejection),
        'reopen_source_preserved': (
            _digest(output / 'created.pptx') == source_hash_before_reopen),
    }
    return {'requested': [width, height], 'checks': checks,
            'pptx_sha256': source_hash_before_reopen,
            'existing_content_rejection': rejection,
            'created': created, 'reopened': snapshot, 'pdf': pdf}


def _generate_recorded_variant(output, *, deadline):
    from skills.WPSComposer.scripts.generation_plan import GenerationOperation, GenerationPlan
    from skills.WPSComposer.scripts.recording_composers import RecordedGeneration
    from skills.WPSComposer.scripts.msoffice.macos_office_runtime import generate_recorded
    plan = GenerationPlan('presentation', (
        GenerationOperation('slide.reset', {}),
        GenerationOperation('slide.set_size', {'width': 720, 'height': 405}),
        GenerationOperation('slide.add_blank', {}),
    ))
    target = output / 'generated-recorded.pptx'
    generated = generate_recorded(RecordedGeneration(plan, ()), target,
                                  timeout=min(90, deadline - time.monotonic()))
    return {'entry': 'generate_recorded', 'requested': [720, 405],
            'pptx_sha256': _digest(generated),
            'size': _pptx_size(generated),
            'passed': _pptx_size(generated) == [720, 405]}


def run(output, *, sentinel, execute=False):
    if execute is not True:
        raise ValueError('Native execution requires execute=True')
    if sys.platform != 'darwin':
        raise RuntimeError('This fixture requires macOS')
    from skills.WPSComposer import create_document, generate, open_document
    from skills.WPSComposer.scripts.msoffice import (
        macos_office_runtime, macos_powerpoint_script, macos_powerpoint_session,
    )
    from skills.WPSComposer.scripts.msoffice.macos_office_runtime import _container_root

    output = Path(output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=False)
    create_dirs = [output / 'public-create-landscape-01',
                   output / 'public-create-portrait-01']
    recorded_dir = output / 'lower-level-recorded-01'
    generate_dir = output / 'public-generate-01'
    for path in create_dirs:
        path.mkdir()
    recorded_dir.mkdir()
    generate_dir.mkdir()
    deadline = time.monotonic() + 600
    container = _container_root('presentation')
    report = {'status': 'RUNNING', 'engine': 'msoffice', 'sentinel': sentinel,
              'checks': {}, 'source_sha256': {
                  'fixture': _digest(__file__),
                  'compiler': _digest(macos_powerpoint_script.__file__),
                  'session': _digest(macos_powerpoint_session.__file__),
                  'runtime': _digest(macos_office_runtime.__file__),
              }}
    session = None
    try:
        before = _inventory(output / 'before.json', container, deadline)
        _check_sentinel(before, sentinel)

        report['create'] = [
            _create_variant(path, width, height, create_document=create_document,
                            open_document=open_document, deadline=deadline)
            for path, (width, height) in zip(create_dirs, [(720, 405), (405, 720)])
        ]
        report['checks']['create_variants'] = all(
            all(variant['checks'].values()) for variant in report['create'])

        report['recorded'] = _generate_recorded_variant(recorded_dir, deadline=deadline)
        report['checks']['lower_level_recorded_size'] = report['recorded']['passed']

        generated = generate(
            '# Public generate default\n\n## Body\n\nGenerated through native PowerPoint.',
            format='pptx', source_is_text=True, output=str(generate_dir / 'generated.pptx'),
            engine='msoffice', timeout=min(120, deadline - time.monotonic()))
        generated_hash_before_reopen = _digest(generated)
        with open_document(generated, kind='slide', engine='msoffice', read_only=True) as reopened:
            reopened.timeout = min(90, deadline - time.monotonic())
            generate_snapshot = reopened.inspect_document()
            reopened.export_pdf(generate_dir / 'generated.pdf')
        _copy_job(reopened, generate_dir / 'reopen-job')
        generated_pdf = _pdf(generate_dir / 'generated.pdf')
        report['checks']['generate_default_size'] = _pptx_size(generated) == [960, 540]
        report['checks']['generate_reopened'] = (
            generate_snapshot['slide_count'] > 0
            and abs(generate_snapshot['page_setup']['slide_width'] - 960) < .1)
        report['checks']['generate_pdf'] = (
            bool(generated_pdf['pages'])
            and all(abs(w - 960) < .1 and abs(h - 540) < .1
                    for w, h in generated_pdf['pages'])
            and 'Public generate default' in generated_pdf['text'])
        report['checks']['generate_source_preserved'] = (
            _digest(generated) == generated_hash_before_reopen)
        report['generate'] = {'pptx_sha256': generated_hash_before_reopen,
                              'snapshot': generate_snapshot, 'pdf': generated_pdf}

        after = _inventory(output / 'after.json', container, deadline)
        report['checks']['sentinel_unchanged'] = before == after
        report['checks']['no_quarantine'] = not (
            container / 'native-office.quarantine.json').exists()
        report['status'] = 'PASS' if all(report['checks'].values()) else 'FAIL'
    except BaseException as exc:
        report['status'] = 'FAILED_RETAINED'
        report['error'] = type(exc).__name__ + ': ' + str(exc)
        raise
    finally:
        (output / 'report.json').write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    return report


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--sentinel', required=True)
    args = parser.parse_args(argv)
    result = run(args.output, sentinel=args.sentinel, execute=args.execute)
    print(json.dumps({'status': result['status'], 'checks': result['checks']},
                     ensure_ascii=False))
    return 0 if result['status'] == 'PASS' else 1


if __name__ == '__main__':
    raise SystemExit(main())
