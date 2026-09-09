"""UNEXECUTED hypothesis: can native orientation swaps expose a height setter?

PowerPoint.sdef declares page setup.slide width and slide orientation, but no
slide height. This probe tries opposite orientation -> width H -> desired
orientation -> width W. It is NOT evidence that orientation swaps dimensions.
Every intermediate save records native width/orientation and read-only OOXML
sldSz diagnostics. A final PDF and native reopen must agree with W/H before a
variant can pass. First failure stops the batch; uncertain owned documents and
their original logs remain quarantined for explicit recovery. No Quit, UI,
clipboard, OOXML writes, or changes to preexisting presentations are used.
Preexisting presentation text is compared only by SHA-256 in persisted records;
raw inventory stdout stays in memory, including partial timeout output.

Prepare only (default): python -m fixtures.microsoft_parity.macos_powerpoint_size_probe
Native execution requires coordinated exclusive UI access and explicit flags:
  --execute --output NEW_DIR --sentinel EXACT_EXISTING_UNSAVED_PRESENTATION_NAME
Use --size WIDTH HEIGHT repeatedly (points); at most three bounded variants.
Fixtures are not installed runtime sources. This report separately records the
fixture, session, dictionary and installed-source hashes; it certifies no parity
catalog row automatically. Importing this module performs no native/file work.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from zipfile import ZipFile


def _dimension(value):
    try:
        valid = type(value) in (int, float) and math.isfinite(value) and value > 0
    except OverflowError:
        valid = False
    if not valid:
        raise ValueError('Dimensions must be positive finite numbers in points')
    return format(value, '.15g')


def compile_steps(width, height):
    width_text, height_text = _dimension(width), _dimension(height)
    desired = 'horizontal orientation' if width >= height else 'vertical orientation'
    opposite = 'vertical orientation' if width >= height else 'horizontal orientation'
    return [
        {'name': 'opposite-orientation', 'command': 'set slide orientation of page setup of ownedDoc to ' + opposite},
        {'name': 'width-as-height', 'command': 'set slide width of page setup of ownedDoc to ' + height_text},
        {'name': 'desired-orientation', 'command': 'set slide orientation of page setup of ownedDoc to ' + desired},
        {'name': 'final-width', 'command': 'set slide width of page setup of ownedDoc to ' + width_text},
    ]


CONTENT_GEOMETRY = {
    'textbox': {'left': 30, 'top': 30, 'width': 260, 'height': 40},
    'rectangle': {'left': 40, 'top': 110, 'width': 180, 'height': 60},
}


def variant_actions(width, height, *, size_before_content=False):
    if type(size_before_content) is not bool:
        raise ValueError('size_before_content must be boolean')
    steps = [dict(step, kind='size') for step in compile_steps(width, height)]
    content = {'kind': 'content', 'name': 'content'}
    return steps + [content] if size_before_content else [content] + steps


def _add_sized_content(session, width, height):
    _, index = session.add_blank_slide()
    text = f'Size probe {width} x {height}'
    target = session.add_textbox(index, text, **CONTENT_GEOMETRY['textbox'])
    # Public creation styling requests fixed textbox bounds explicitly. Native
    # text boxes otherwise default to auto height, independently of slide size.
    session.apply_format_patch(target, text_frame={'auto_size': 0},
                               geometry=CONTENT_GEOMETRY['textbox'])
    rectangle = session.add_shape(index, 1, **CONTENT_GEOMETRY['rectangle'],
                                  fill_color='#2463A6', text='Native rectangle')
    return {'textbox': {'target': target, 'text': text},
            'rectangle': {'target': rectangle, 'text': 'Native rectangle'}}


def _content_matches(snapshot, contract):
    shapes = [shape for slide in snapshot.get('slides', []) for shape in slide.get('shapes', [])]
    for role, expected in contract.items():
        found = [shape for shape in shapes if shape.get('text', '').strip() == expected['text']]
        if len(found) != 1:
            return False
        if any(abs(found[0]['geometry'][key] - value) >= 0.1
               for key, value in CONTENT_GEOMETRY[role].items()):
            return False
    return True


def require_owned(session, container):
    job = Path(str(session._job)).resolve()
    path = Path(str(session._path)).resolve()
    container = Path(container).resolve()
    if (not session._entered or session._attached or session._uncertain
            or job.parent != container or not job.name.startswith('session-')
            or path.parent != job or path.suffix != '.pptx'):
        raise RuntimeError('Probe requires an entered, certain, private owned PowerPoint session')


def retain_failure(session, error):
    """No more AppleEvents after failure, including an automatic context close."""
    try:
        if session._entered and not session._uncertain:
            session._quarantine('Size probe stopped: ' + type(error).__name__)
    finally:
        session._entered = False
        session._release()


def _digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_size(path):
    with ZipFile(path, 'r') as package:
        root = ET.fromstring(package.read('ppt/presentation.xml'))
    size = root.find('{http://schemas.openxmlformats.org/presentationml/2006/main}sldSz')
    if size is None:
        raise ValueError('Native artifact has no slide size')
    return {'width_pt': int(size.attrib['cx']) / 12700,
            'height_pt': int(size.attrib['cy']) / 12700}


def _write(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n')


def _redact_inventory(rows):
    if not isinstance(rows, list):
        raise ValueError('Unexpected inventory response')
    redacted = []
    for row in rows:
        if not isinstance(row, list) or len(row) != 7 or not isinstance(row[4], list):
            raise ValueError('Unexpected inventory row')
        content = json.dumps(row[4], ensure_ascii=False, separators=(',', ':')).encode('utf-8')
        redacted.append(row[:4] + [{'sha256': hashlib.sha256(content).hexdigest(),
                                    'has_text': any(bool(text) for text in row[4])}] + row[5:])
    return sorted(redacted, key=lambda row: (row[0], row[1]))


def _stdout_metadata(output, value):
    content = value if isinstance(value, bytes) else (value or '').encode('utf-8')
    _write(output.with_suffix('.stdout-metadata.json'),
           {'bytes': len(content), 'sha256': hashlib.sha256(content).hexdigest(),
            'raw_stdout_retained': False})


def _inventory(output, container, deadline):
    """Read-only snapshot under the same app lock; any error halts the probe."""
    from skills.WPSComposer.scripts.msoffice.macos_office_runtime import OfficeJobLock
    from skills.WPSComposer.scripts.msoffice.macos_powerpoint_session import _JSON
    script = _JSON + '''if application "Microsoft PowerPoint" is not running then error "PowerPoint must already be running with the requested sentinel"
with timeout of 20 seconds
 tell application "Microsoft PowerPoint"
  set inventoryRows to {}
  repeat with di from 1 to count of presentations
   set docRef to presentation di
   set allText to {}
   repeat with si from 1 to count of slides of docRef
    repeat with sh from 1 to count of shapes of slide si of docRef
     if has text frame of shape sh of slide si of docRef then set end of allText to content of text range of text frame of shape sh of slide si of docRef
    end repeat
   end repeat
   set end of inventoryRows to {name of docRef,full name of docRef,saved of docRef,count of slides of docRef,allText,slide width of page setup of docRef,slide orientation of page setup of docRef as text}
  end repeat
  return my encodeJSON(inventoryRows)
 end tell
end timeout'''
    container.mkdir(parents=True, exist_ok=True)
    lock = OfficeJobLock(container)
    lock.acquire(deadline)
    source = output.with_suffix('.applescript')
    source.write_text(script)
    try:
        remaining = min(25, deadline - time.monotonic())
        if remaining <= 0:
            raise TimeoutError('Probe deadline expired')
        result = subprocess.run(['/usr/bin/osascript', str(source)], capture_output=True,
                                text=True, timeout=remaining)
        _stdout_metadata(output, result.stdout)
        output.with_suffix('.stderr').write_text(result.stderr)
        if result.returncode:
            raise RuntimeError('Native inventory failed; raw logs retained')
        rows = _redact_inventory(json.loads(result.stdout))
        _write(output.with_suffix('.stdout'), rows)
        _write(output, rows)
        return rows
    except BaseException as exc:
        if isinstance(exc, subprocess.TimeoutExpired):
            _stdout_metadata(output, exc.stdout)
            value = exc.stderr
            output.with_suffix('.stderr').write_bytes(value if isinstance(value, bytes) else (value or '').encode())
        lock.quarantine({'component': 'presentation', 'reason': 'Size probe inventory failed',
                         'diagnostic_path': str(source)})
        raise
    finally:
        lock.close()


def _native_size(session):
    return json.loads(session._run('return my encodeJSON({slide width of page setup of ownedDoc, slide orientation of page setup of ownedDoc as text})'))


def _deadline_session(session, deadline):
    session.timeout = min(60, deadline - time.monotonic())
    if session.timeout <= 0:
        raise TimeoutError('Probe deadline expired')
    session.__enter__()


def _add_probe_slide(session, width, height):
    _, slide_index = session.add_blank_slide()
    session.add_textbox(slide_index, f'Size probe {width} x {height}', 30, 30, 260, 40)


def run(output, sizes, *, sentinel_names, execute=False, size_before_content=False):
    if execute is not True:
        raise ValueError('Native execution requires execute=True')
    plans = [(width, height, variant_actions(width, height, size_before_content=size_before_content)) for width, height in sizes]
    if not 1 <= len(plans) <= 3:
        raise ValueError('Use one to three size variants')
    if (not sentinel_names or len(set(sentinel_names)) != len(sentinel_names)
            or any(not isinstance(name, str) or not name for name in sentinel_names)):
        raise ValueError('Name at least one existing unsaved synthetic sentinel')
    if sys.platform != 'darwin':
        raise RuntimeError('This guarded probe requires macOS')
    from skills.WPSComposer import create_document, open_document
    from skills.WPSComposer.scripts.msoffice.macos_office_runtime import _container_root
    from skills.WPSComposer.scripts.msoffice import macos_powerpoint_session
    from fixtures.microsoft_parity.evidence_gate import source_digest
    from pypdf import PdfReader

    output = Path(output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=False)
    container = _container_root('presentation')
    deadline = time.monotonic() + 600
    report = {'status': 'RUNNING', 'hypothesis_only': True, 'engine': 'msoffice',
              'platform': sys.platform, 'variants': [], 'native_jobs': [],
              'fixture_sha256': _digest(__file__),
              'session_sha256': _digest(macos_powerpoint_session.__file__),
              'source_digest': source_digest(Path(__file__).resolve().parents[2]),
              'sentinel_names': sentinel_names, 'size_before_content': size_before_content}
    dictionary = Path('/Applications/Microsoft PowerPoint.app/Contents/Resources/PowerPoint.sdef')
    report['dictionary_sha256'] = _digest(dictionary)
    session = None
    try:
        baseline = _inventory(output / 'before.json', container, deadline)
        for name in sentinel_names:
            matches = [row for row in baseline if row[0] == name]
            if len(matches) != 1 or matches[0][2] is not False or not matches[0][4]['has_text']:
                raise ValueError('Requested sentinel must uniquely identify an unsaved presentation with text')
        for number, (width, height, steps) in enumerate(plans, 1):
            variant_dir = output / f'variant-{number:02}'
            variant_dir.mkdir()
            variant = {'requested': {'width_pt': width, 'height_pt': height}, 'steps': []}
            report['variants'].append(variant)
            session = create_document('slide', engine='msoffice', visible=False)
            _deadline_session(session, deadline)
            report['native_jobs'].append(str(session._job))
            require_owned(session, container)
            content_contract = None
            if size_before_content:
                session.save(variant_dir / '00-empty-baseline.pptx')
            size_index = 0
            for step in steps:
                require_owned(session, container)
                if step['kind'] == 'content':
                    if size_before_content:
                        content_contract = _add_sized_content(session, width, height)
                        variant['content_contract'] = content_contract
                        variant['created_snapshot'] = session.inspect_document()
                        path = variant_dir / '05-created-content.pptx'
                    else:
                        _add_probe_slide(session, width, height)
                        path = variant_dir / '00-baseline.pptx'
                    session.save(path)
                    continue
                size_index += 1
                session._run(step['command'] + '\nreturn "SIZE_STEP_OK"', mutation=True)
                path = variant_dir / f'{size_index:02}-{step["name"]}.pptx'
                session.save(path)
                variant['steps'].append({**step, 'native': _native_size(session),
                                         'artifact': str(path), 'sha256': _digest(path),
                                         'ooxml': read_size(path)})
                _write(output / 'report.json', report)
            final_path = path
            session.export_pdf(variant_dir / 'final.pdf')
            if not session.is_bound_to(session._path):
                raise RuntimeError('Owned identity changed after PDF export')
            session.close(save_changes=False)
            session = None
            final_hash = _digest(final_path)
            session = open_document(final_path, kind='slide', engine='msoffice', read_only=True)
            _deadline_session(session, deadline)
            report['native_jobs'].append(str(session._job))
            require_owned(session, container)
            variant['reopened_native'] = _native_size(session)
            variant['reopened_ooxml'] = read_size(session._path)
            if size_before_content:
                variant['reopened_snapshot'] = session.inspect_document()
            session.close(save_changes=False)
            session = None
            variant['reopen_source_preserved'] = _digest(final_path) == final_hash
            pdf = PdfReader(variant_dir / 'final.pdf')
            variant['pdf_pages'] = [[float(page.mediabox.width), float(page.mediabox.height)] for page in pdf.pages]
            if size_before_content:
                variant['pdf_text'] = '\n'.join(page.extract_text() or '' for page in pdf.pages)
                variant['content_confirmed'] = (
                    _content_matches(variant['created_snapshot'], content_contract) and
                    _content_matches(variant['reopened_snapshot'], content_contract) and
                    all(item['text'] in variant['pdf_text'] for item in content_contract.values()))
            actual = variant['reopened_ooxml']
            variant['size_confirmed'] = (abs(actual['width_pt'] - width) < 0.1 and
                abs(actual['height_pt'] - height) < 0.1 and
                abs(float(variant['reopened_native'][0]) - width) < 0.1 and variant['reopen_source_preserved'] and
                bool(variant['pdf_pages']) and all(abs(w - width) < 0.1 and abs(h - height) < 0.1
                                                  for w, h in variant['pdf_pages']))
            after = _inventory(variant_dir / 'after.json', container, deadline)
            variant['preexisting_documents_unchanged'] = baseline == after
            if not variant['preexisting_documents_unchanged']:
                raise RuntimeError('Preexisting presentation snapshot changed; stop without recovery actions')
            if not variant['size_confirmed']:
                report['status'] = 'HYPOTHESIS_NOT_CONFIRMED'
                break
            if size_before_content and not variant['content_confirmed']:
                report['status'] = 'CONTENT_NOT_CONFIRMED'
                break
        else:
            report['status'] = 'TESTED_SIZE_BEFORE_CONTENT_CONFIRMED' if size_before_content else 'TESTED_SIZES_CONFIRMED'
    except BaseException as exc:
        report.update(status='FAILED_RETAINED', error=type(exc).__name__ + ': ' + str(exc))
        if session is not None:
            report['retained_job'] = str(session._job)
            retain_failure(session, exc)
        raise
    finally:
        _write(output / 'report.json', report)
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', type=Path)
    parser.add_argument('--sentinel', action='append', default=[])
    parser.add_argument('--size', type=float, nargs=2, action='append')
    parser.add_argument('--size-before-content', action='store_true')
    args = parser.parse_args(argv)
    sizes = args.size or [(720, 405), (405, 720)]
    if not args.execute:
        print(json.dumps({'status': 'PREPARED_NOT_EXECUTED', 'hypothesis_only': True,
                          'variants': [variant_actions(*size, size_before_content=args.size_before_content) for size in sizes]}, indent=2))
        return 0
    if args.output is None:
        parser.error('--execute requires --output NEW_DIR')
    result = run(args.output, sizes, sentinel_names=args.sentinel, execute=True,
                 size_before_content=args.size_before_content)
    return 0 if result['status'] in ('TESTED_SIZES_CONFIRMED', 'TESTED_SIZE_BEFORE_CONTENT_CONFIRMED') else 1


if __name__ == '__main__':
    raise SystemExit(main())
