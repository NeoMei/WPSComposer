"""Explicit, bounded Windows Microsoft public API acceptance.

Run with the selected Python environment and --plugin INSTALL or --checkout ROOT,
and a NEW --output directory. Each app runs in its own deadline-bound Python
child. This runner never kills Office or closes an unknown document. A timeout
retains logs and quarantine; it is not evidence of native cleanup.

--fault-timeout is opt-in and runs AFTER normal acceptance. It intentionally
leaves one owned Office document and its recovery files per selected component;
the production quarantine remains until separately verified manual recovery.
No native call occurs during import, --help, or the portable test suite.
"""
from __future__ import annotations

import argparse
import contextlib
import hashlib
import importlib.util
import json
import math
import ntpath
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from uuid import uuid4

KINDS = ('writer', 'sheet', 'slide')
SUFFIX = dict(writer='docx', sheet='xlsx', slide='pptx')
COMPONENT = dict(writer='writer', sheet='spreadsheet', slide='presentation')
EXECUTABLE = dict(writer='WINWORD.EXE', sheet='EXCEL.EXE', slide='POWERPNT.EXE')
MARKER = dict(writer='Native Word', sheet='Native Excel', slide='Native PowerPoint')
REQUIRED = {'public_create_save_reopen', 'public_edit_reopen', 'public_generate',
            'public_convert_pdf', 'pdf_contents', 'typed_handles', 'source_preserved',
            'proxy_identity_close', 'sentinels_preserved', 'baseline_preserved',
            'sentinels_closed', 'source_unchanged'}


def write_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')


def digest(path):
    result = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            result.update(block)
    return result.hexdigest()


def retain_quarantine(root, detail):
    root = Path(root)
    if root.is_symlink(): raise ValueError('quarantine root must not be a symlink')
    root.mkdir(parents=True, exist_ok=True)
    marker = root / 'native-office.quarantine.json'
    try:
        with marker.open('x', encoding='utf-8') as stream:
            json.dump(dict(detail, cleanup_verified=False, office_termination_attempted=False), stream)
    except FileExistsError:
        pass  # Preserve the first recovery record, including production details.


def run_bounded(command, output, *, timeout):
    """Only terminate the exact Popen handle; never a process name or tree."""
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    with (output / 'stdout.log').open('wb') as stdout, (output / 'stderr.log').open('wb') as stderr:
        child = subprocess.Popen(command, stdout=stdout, stderr=stderr, stdin=subprocess.DEVNULL)
        report = {'status':'completed', 'python_pid':child.pid, 'office_termination_attempted':False}
        try:
            child.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            report.update(status='timeout', recovery_required=True,
                          cleanup_verified=False, command=command)
            # The outer deadline can interrupt sentinel COM. No Office object
            # is assumed closed, and nested workers are not killed by PID.
            child.kill()
            try:
                child.wait(timeout=5)
            except subprocess.TimeoutExpired:
                report['python_exit_unverified'] = True
            write_json(output / 'outer-timeout.json', report)
            manifest = output / 'recovery-root.json'
            if manifest.is_file():
                try:
                    root = json.loads(manifest.read_text(encoding='utf-8'))['root']
                    retain_quarantine(root, dict(report, evidence_path=str(output)))
                except (OSError, ValueError, KeyError) as exc:
                    report['quarantine_error'] = str(exc)
        report['returncode'] = child.poll()
    return report


def verify_proxy_runtime(job, kind, *, returncode):
    job = Path(job)
    rows = [json.loads(line) for line in (job / 'worker-stdout.log').read_text(encoding='utf-8').splitlines()]
    assert rows and returncode == 0, 'proxy did not exit normally'
    identities = [r.get('identity') for r in rows if r.get('status') == 'ok']
    assert identities and all(i == identities[0] for i in identities), 'proxy Office identity changed'
    identity = identities[0]
    assert type(identity.get('office_pid')) is int and identity['office_pid'] > 0, 'missing Office identity PID'
    assert ntpath.basename(identity.get('office_executable', '')).upper() == EXECUTABLE[kind], 'wrong Office identity executable'
    assert identity.get('attached') is False, 'unexpected attached identity'
    stage = Path(identity.get('staging_path', '')).resolve()
    assert job.resolve() in stage.parents, 'native staging identity escaped proxy job'
    request = json.loads((job / 'last-request.json').read_text(encoding='utf-8'))
    last = rows[-1]
    assert (request['method'] == 'close' and last.get('id') == request['id'] and
            last.get('protocol') == 1 and last.get('status') == 'ok' and
            last.get('value') == {'closed':True}), 'missing exact final close ACK'
    assert all(r.get('protocol') == 1 and r.get('id') == n for n, r in enumerate(rows, 1)), 'invalid response sequence'
    return {'identity':identity, 'close_ack':True, 'python_returncode':returncode}


def check_snapshot(kind, snapshot, *, edited=False):
    """Assert actual schema locations, not a substring in a JSON dump."""
    marker = MARKER[kind] + (' edited' if edited and kind != 'sheet' else '')
    if kind == 'writer':
        assert snapshot['paragraphs'][0]['text'] == marker
        assert len(snapshot['tables']) == 1
        assert any(c.get('text', '').rstrip('\r\x07') == 'Native Table' for c in snapshot['tables'][0]['cells'])
        assert snapshot['counts']['shapes'] == 0, 'removed typed shape persisted'
    elif kind == 'sheet':
        data = [s for s in snapshot['sheets'] if s['name'] == 'Data']
        details = [s for s in snapshot['sheets'] if s['name'] == 'Details']
        assert len(data) == len(details) == 1 and data[0]['index'] == 1
        cells = {c['address'].replace('$', ''):c for c in data[0]['cells']}
        assert {'A1', 'A2', 'B2'} <= cells.keys(), 'required exact Excel cells are missing'
        assert cells['A1']['value'] == marker
        assert cells['A2']['value'] == (5 if edited else 4)
        assert cells['B2']['formula'] == '=A2*3'
        assert cells['B2']['value'] == (15 if edited else 12)
        assert any(c['address'].replace('$', '') == 'A1' and c['value'] == 'Typed worksheet' for c in details[0]['cells'])
    else:
        slides = snapshot['slides']
        assert len(slides) == 2 and slides[0]['index'] == 1 and slides[1]['index'] == 2
        assert [s['text'] for s in slides[0]['shapes']] == [marker]
        assert [s['text'] for s in slides[1]['shapes']] == ['Typed slide']
    return True


def assert_baseline(before, after):
    assert before == after, 'native document baseline changed'


def check_generated_snapshot(kind, snapshot):
    expected = {(1,1):'Item', (1,2):'Value', (2,1):'Native row', (2,2):'42'}
    if kind == 'sheet':
        # The sheet renderer intentionally emits tables, with the source
        # heading as sheet name; standalone prose is not an XLSX capability.
        sheets = [s for s in snapshot['sheets'] if s['name'] == 'Generated Native']
        assert len(sheets) == 1, 'generated worksheet heading missing'
        cells = {c['address'].replace('$', ''):c['value'] for c in sheets[0]['cells']}
        assert cells.get('A1') == 'Item' and cells.get('B1') == 'Value'
        assert cells.get('A2') == 'Native row' and cells.get('B2') in ('42', 42, 42.0)
    else:
        if kind == 'writer':
            texts = [p.get('text', '').rstrip('\r\x07') for p in snapshot['paragraphs']]
            tables = snapshot['tables']
        else:
            shapes = [shape for slide in snapshot['slides'] for shape in slide['shapes']]
            texts = [s.get('text', '').rstrip('\r\n') for s in shapes]
            tables = [s['table'] for s in shapes if 'table' in s]
        assert 'Generated Native' in texts, 'generated title missing'
        assert 'Generated native content.' in texts, 'generated paragraph missing'
        assert len(tables) == 1, 'generated native table missing or duplicated'
        actual = {(c['row'], c['column']):c.get('text', '').rstrip('\r\x07') for c in tables[0]['cells']}
        assert actual == expected, 'generated native table values/positions differ from source'
    return True


def pdf_check(path, required):
    from pypdf import PdfReader
    reader = PdfReader(path)
    assert reader.pages, 'empty PDF'
    text = '\n'.join(page.extract_text() or '' for page in reader.pages)
    for marker in required:
        assert marker in text, f'PDF missing {marker!r}'
    return {'pages':len(reader.pages), 'text':text, 'sha256':digest(path)}


def import_api(plugin):
    plugin = Path(plugin).resolve(strict=True)
    sys.path.insert(0, str(plugin))
    import skills.WPSComposer as api
    assert plugin in Path(api.__file__).resolve().parents, 'public API imported outside selected bundle'
    return api


def bound_source_digest(plugin):
    source = Path(__file__).with_name('evidence_gate.py')
    spec = importlib.util.spec_from_file_location('windows_runner_evidence_gate', source)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.source_digest(plugin)


def _text_content(doc, kind):
    if kind == 'writer':
        return str(doc.Content.Text)
    if kind == 'sheet':
        return [(str(doc.Worksheets.Item(i).Name), doc.Worksheets.Item(i).UsedRange.Value,
                 doc.Worksheets.Item(i).UsedRange.Formula) for i in range(1, int(doc.Worksheets.Count) + 1)]
    return [[str(s.TextFrame.TextRange.Text) if s.HasTextFrame else ''
             for s in (doc.Slides.Item(i).Shapes.Item(j) for j in range(1, int(doc.Slides.Item(i).Shapes.Count) + 1))]
            for i in range(1, int(doc.Slides.Count) + 1)]


class NativeSentinels:
    """Own exact references; baseline documents are read-only throughout."""
    def __init__(self, kind, output):
        self.kind, self.output = kind, Path(output)
        self.composer = None
        self.owned = []
        self.apps = []
        self.initial_registered = None
        self.tokens = []  # Keep IUnknown references alive; repr alone is not identity.

    def _token(self, obj):
        value = self.deps.identity(obj)
        for index, token in enumerate(self.tokens):
            if token == value:
                return index
        self.tokens.append(value)
        return len(self.tokens) - 1

    def _collection(self, app):
        return getattr(app, {'writer':'Documents', 'sheet':'Workbooks', 'slide':'Presentations'}[self.kind])

    def _record(self, doc):
        text = json.dumps(_text_content(doc, self.kind), ensure_ascii=False, default=str)
        if self.kind == 'slide':
            windows = [hwnd for hwnd, window in self.deps.powerpoint_windows()
                       if self.deps.identity(window.Application) == self.deps.identity(doc.Application)
                       and window.Presentation is not None and self.deps.identity(window.Presentation) == self.deps.identity(doc)]
        else:
            windows = [int(doc.Windows.Item(i).Hwnd) for i in range(1, int(doc.Windows.Count) + 1)]
        bindings = [{'hwnd':hwnd, 'pid':self.deps.window_pid(hwnd),
                     'executable':self.deps.process_image(self.deps.window_pid(hwnd))} for hwnd in windows]
        return {'identity':self._token(doc), 'name':str(doc.Name), 'full_name':str(doc.FullName),
                'application_identity':self._token(doc.Application), 'windows':bindings,
                'saved':bool(doc.Saved), 'content_sha256':hashlib.sha256(text.encode()).hexdigest()}

    def snapshot(self, app):
        collection = self._collection(app)
        return [self._record(collection.Item(i)) for i in range(1, int(collection.Count) + 1)]

    def start(self):
        from skills.WPSComposer.scripts.msoffice import windows_host, windows_office_host
        import pythoncom
        import win32com.client
        self.deps = windows_host._load_dependencies() if self.kind == 'writer' else windows_office_host._load_dependencies(COMPONENT[self.kind])
        # The factories initialize COM too. This explicit apartment covers the
        # pre-existing registered application read, before factory creation.
        pythoncom.CoInitialize()
        self.pythoncom = pythoncom
        self.before_processes = self.deps.processes()
        progid = {'writer':'Word.Application', 'sheet':'Excel.Application', 'slide':'PowerPoint.Application'}[self.kind]
        try:
            registered = win32com.client.GetActiveObject(progid)
        except pythoncom.com_error as exc:
            if exc.hresult != -2147221021:  # MK_E_UNAVAILABLE only.
                raise
            registered = None
        if registered is not None:
            self.initial_registered = (registered, self.snapshot(registered))
            self.apps.append(registered)
        root = self.output / 'sentinel-private'
        root.mkdir()
        self.composer = (windows_host.create_dedicated_composer(str(root)) if self.kind == 'writer'
                         else windows_office_host.create_composer(COMPONENT[self.kind], root))
        self.composer._verify_application()
        app = self.composer._app
        if not any(self._token(app) == self._token(old) for old in self.apps):
            self.apps.append(app)
        primary = self.composer._doc
        self.owned.append((primary, self._token(primary)))
        self._set_marker(primary, 'UNSAVED-SENTINEL-' + uuid4().hex)
        assert not bool(primary.Saved), 'sentinel is not unsaved'
        self.composer._verify_application()
        saved = self._collection(app).Add()
        self.owned.append((saved, self._token(saved)))
        assert self._token(saved.Application) == self._token(app)
        self._set_marker(saved, 'SAVED-SENTINEL-' + uuid4().hex)
        path = root / ('saved-' + uuid4().hex + '.' + SUFFIX[self.kind])
        if self.kind == 'writer': saved.SaveAs2(str(path), FileFormat=16, AddToRecentFiles=False)
        elif self.kind == 'sheet': saved.SaveAs(str(path), FileFormat=51)
        else: saved.SaveAs(str(path), 24)
        assert ntpath.normcase(str(saved.FullName)) == ntpath.normcase(str(path)) and bool(saved.Saved)
        self.saved_path, self.saved_digest = path, digest(path)
        self.expected = [self._record(doc) for doc, _ in self.owned]
        assert all(d['windows'] and all(ntpath.basename(w['executable'] or '').upper() == EXECUTABLE[self.kind]
                   for w in d['windows']) for d in self.expected), 'sentinel native HWND identity missing'
        self.baseline = [self.snapshot(app) for app in self.apps]
        write_json(self.output / 'sentinels-before.json', {'documents':self.expected,
                   'application_baselines':self.baseline, 'processes':self.before_processes})

    def _set_marker(self, doc, marker):
        if self.kind == 'writer': doc.Content.Text = marker
        elif self.kind == 'sheet': doc.Worksheets.Item(1).Cells(1, 1).Value = marker
        else:
            slide = doc.Slides.Add(1, 12)
            slide.Shapes.AddTextbox(1, 40, 40, 500, 80).TextFrame.TextRange.Text = marker

    def verify(self, *, allow_owned_extra=False):
        self.composer._verify_application()
        observed = [self._record(doc) for doc, _ in self.owned]
        assert_baseline(self.expected, observed)
        assert digest(self.saved_path) == self.saved_digest, 'saved sentinel file changed'
        after = [self.snapshot(app) for app in self.apps]
        if not allow_owned_extra:
            assert_baseline(self.baseline, after)
        # Fault runs may intentionally retain a NEW worker-owned document in
        # shared PowerPoint. Every original document still must match exactly.
        else:
            for old, new in zip(self.baseline, after):
                ids = {d['identity']: d for d in new}
                assert_baseline(old, [ids.get(d['identity']) for d in old])
        for pid, image in self.before_processes.items():
            assert self.deps.process_image(pid) == image, 'baseline Office process changed'
        write_json(self.output / 'sentinels-after.json', {'documents':observed, 'application_baselines':after})

    def close(self, *, preserve_host=False):
        # An interrupted setup also retains exact references, but only fully
        # fingerprinted sentinels are eligible for automatic discard.
        if not hasattr(self, 'expected'):
            return False
        for (doc, token), expected in reversed(list(zip(self.owned, self.expected))):
            self.composer._verify_application()
            assert self._token(doc) == token and self._record(doc) == expected, 'sentinel changed; preserved for recovery'
            if self.kind == 'slide':
                doc.Saved = True
                doc.Close()
            else:
                doc.Close(SaveChanges=False)
        owned_tokens = {token for _, token in self.owned}
        for app in self.apps:
            assert not any(row['identity'] in owned_tokens for row in self.snapshot(app)), 'owned sentinel is still open after Close'
        # References are dropped only after exact collection absence is proven.
        self.composer._doc = None
        if preserve_host:
            # Fault worker documents may share this application. Discard only
            # exact sentinel references; retain host and unknown worker state.
            self.composer._app = None
            if self.composer._com_initialized:
                self.deps.uninitialize()
                self.composer._com_initialized = False
        else:
            self.composer.close(save_changes=False)
        if self.initial_registered is not None:
            app, before = self.initial_registered
            # Factories cannot own a registered baseline process; its complete
            # document collection must still be present after sentinel cleanup.
            after = self.snapshot(app)
            if preserve_host:
                indexed = {row['identity']:row for row in after}
                assert_baseline(before, [indexed.get(row['identity']) for row in before])
            else:
                assert_baseline(before, after)
        self.tokens.clear()
        self.pythoncom.CoUninitialize()
        return True


@contextlib.contextmanager
def observe_sessions(output, deadline):
    """Observe actual public factories, including inspect/edit hidden sessions."""
    from skills.WPSComposer.scripts.msoffice.windows_session_proxy import _SessionProxy
    original = _SessionProxy.__dict__['_start']
    sessions = []
    def start(cls, method, args, kwargs):
        session = original.__func__(cls, method, args, kwargs)
        session._deadline = min(session._deadline, deadline)
        sessions.append(session)
        write_json(Path(output) / 'observed-jobs.json', [{'kind':s.kind, 'job':str(s.staging_root),
            'python_pid':s._child.pid, 'identity':s._identity} for s in sessions])
        return session
    _SessionProxy._start = classmethod(start)
    try:
        yield sessions
    finally:
        _SessionProxy._start = original
        evidence = []
        for index, session in enumerate(sessions, 1):
            job = session.staging_root
            # Copy only AFTER native close or failure, including final ACK.
            copied = Path(output) / 'runtime' / str(index)
            if job.is_dir(): shutil.copytree(job, copied)
            record = {'job':str(job), 'kind':session.kind, 'python_pid':session._child.pid,
                      'returncode':session._child.poll(), 'closed':session._closed,
                      'uncertain':session._uncertain, 'identity':session._identity}
            if session._closed:
                record['verification'] = verify_proxy_runtime(job, session.kind, returncode=session._child.poll())
            evidence.append(record)
        write_json(Path(output) / 'proxy-evidence.json', evidence)


def populate(session, kind):
    from skills.WPSComposer.scripts.msoffice.windows_session_proxy import NativeSessionHandle
    if kind == 'writer':
        session.add_paragraph(MARKER[kind])
        table = session.add_table(2, 2, [['Header', 'Value'], ['Native Table', '42']])
        assert isinstance(table, NativeSessionHandle) and table.type == 'table'
        session.apply_format_patch('table:1/cell:2,1', font={'size':11})
        shape = session.add_floating_textbox('Temporary typed shape', left=40, top=400, width=150, height=35)
        assert isinstance(shape, NativeSessionHandle) and shape.type == 'shape'
        session.apply_structural_op({'op':'remove', 'target':shape})
    elif kind == 'sheet':
        first = session.select_sheet(1)
        assert isinstance(first, NativeSessionHandle) and first.type == 'sheet'
        session.rename_sheet(first, 'Data')
        other = session.add_sheet('Details')
        session.write_cell(1, 1, 'Typed worksheet')
        session.apply_structural_op({'op':'move', 'target':other, 'to':{'before':first}})
        session.apply_structural_op({'op':'move', 'target':other, 'to':{'after':first}})
        session.select_sheet(first)
        session.write_cell(1, 1, MARKER[kind])
        session.write_cell(2, 1, 4)
        session.set_formula(2, 2, '=A2*3')
    else:
        first, _ = session.add_blank_slide()
        second, _ = session.add_blank_slide()
        assert isinstance(first, NativeSessionHandle) and first.type == 'slide'
        shape = session.add_textbox(first, MARKER[kind], 40, 40, 600, 80)
        session.add_textbox(second, 'Typed slide', 40, 40, 600, 80)
        session.apply_structural_op({'op':'move', 'target':first, 'to':{'after':second}})
        session.apply_structural_op({'op':'move', 'target':first, 'to':{'before':second}})
        result = session.apply_structural_op({'op':'move', 'target':shape, 'to':{'slide':first}})
        assert result['moved'] is True and result['to_slide'] == 1
        session.apply_format_patch(shape, font={'size':20})


def standard_case(api, output, kind, deadline):
    output = Path(output)
    checks = {}
    source, edited = output / ('created.' + SUFFIX[kind]), output / ('edited.' + SUFFIX[kind])
    with observe_sessions(output, deadline) as sessions:
        with api.create_document(kind, engine='msoffice') as session:
            populate(session, kind)
            checks['typed_handles'] = True
            session.save(source)
            assert session.is_bound_to(source), 'save changed exact source binding'
            session.save_copy(output / ('copy.' + SUFFIX[kind]))
            session.export_pdf(output / 'created.pdf')
        original = digest(source)
        snapshot = api.inspect(source, engine='msoffice')
        write_json(output / 'created-snapshot.json', snapshot)
        checks['public_create_save_reopen'] = check_snapshot(kind, snapshot)
        copied = api.inspect(output / ('copy.' + SUFFIX[kind]), engine='msoffice')
        check_snapshot(kind, copied)
        target = {'writer':'paragraph:1', 'sheet':'sheet:1/cell:A2', 'slide':'slide:1/shape:1'}[kind]
        patch = {'value':5} if kind == 'sheet' else {'text':MARKER[kind] + ' edited'}
        result = api.edit(source, patches=[{'target':target, **patch}], output=edited,
                          export_pdf=output / 'edited.pdf', engine='msoffice')
        write_json(output / 'edit-result.json', result)
        assert result['ok'] and result['saved'], 'public edit failed'
        reopened = api.inspect(edited, engine='msoffice')
        write_json(output / 'edited-snapshot.json', reopened)
        checks['public_edit_reopen'] = check_snapshot(kind, reopened, edited=True)
        checks['source_preserved'] = digest(source) == original
        generated = output / ('generated.' + SUFFIX[kind])
        markdown = '# Generated Native\n\nGenerated native content.\n\n| Item | Value |\n| --- | --- |\n| Native row | 42 |\n'
        (output / 'source.md').write_text(markdown, encoding='utf-8')
        path = api.generate(markdown, source_is_text=True, format=SUFFIX[kind], output=str(generated),
                            engine='msoffice', timeout=min(600, deadline - time.monotonic()), open_result=False)
        assert Path(path).resolve() == generated.resolve() and generated.is_file()
        generated_snapshot = api.inspect(generated, engine='msoffice')
        write_json(output / 'generated-snapshot.json', generated_snapshot)
        checks['public_generate'] = check_generated_snapshot(kind, generated_snapshot)
        converted = output / 'converted.pdf'
        generated_hash = digest(generated)
        path = api.convert_to_pdf(str(generated), output=str(converted), engine='msoffice',
                                 timeout=min(600, deadline - time.monotonic()), open_result=False)
        assert Path(path).resolve() == converted.resolve() and digest(generated) == generated_hash
        checks['public_convert_pdf'] = True
        pdfs = {'created':pdf_check(output / 'created.pdf', [MARKER[kind]]),
                'edited':pdf_check(output / 'edited.pdf', [MARKER[kind] + (' edited' if kind != 'sheet' else '')]),
                'converted':pdf_check(converted, ['Item', 'Value', 'Native row', '42'] +
                                     ([] if kind == 'sheet' else ['Generated Native', 'Generated native content.']))}
        write_json(output / 'pdf-evidence.json', pdfs)
        checks['pdf_contents'] = True
    assert sessions and all(s._closed and not s._uncertain for s in sessions)
    checks['proxy_identity_close'] = True
    return checks


def fault_case(api, output, kind):
    """Inject a Python-only stall AFTER successful exact native document open."""
    from skills.WPSComposer.scripts.msoffice import windows_session_proxy as proxy
    original = proxy._spawn_worker
    def spawn(job):
        env = os.environ.copy()
        env['PYTHONPATH'] = str(proxy._PACKAGE_ROOT)
        for key in ('TMP', 'TEMP', 'TMPDIR'): env[key] = str(job)
        command = [sys.executable, '-u', str(Path(__file__).resolve()), '--delayed-worker', str(job),
                   '--plugin', str(proxy._PACKAGE_ROOT)]
        with (job / 'worker-stderr.log').open('wb') as log:
            return subprocess.Popen(command, env=env, cwd=str(proxy._PACKAGE_ROOT),
                                    stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=log)
    proxy._spawn_worker = spawn
    session = None
    try:
        session = api.create_document(kind, engine='msoffice')
        identity = dict(session._identity)
        assert identity.get('office_pid') and ntpath.basename(identity.get('office_executable', '')).upper() == EXECUTABLE[kind]
        session._deadline = min(session._deadline, time.monotonic() + 2)
        started = time.monotonic()
        try:
            session.inspect_document()
        except Exception as exc:
            assert str(getattr(exc, 'code', '')).endswith('_TIMEOUT'), repr(exc)
        else:
            raise AssertionError('injected native worker stall unexpectedly completed')
        assert time.monotonic() - started < 10 and session._child.poll() is not None
        recovery = json.loads((session.staging_root / 'recovery.json').read_text())
        assert recovery['python_pid'] == session._child.pid and recovery['identity'] == identity
        assert recovery['office_termination_attempted'] is False and recovery['cleanup_verified'] is False
        assert session._lock.quarantine_path.is_file()
        from skills.WPSComposer.scripts.msoffice import windows_host, windows_office_host
        deps = windows_host._load_dependencies() if kind == 'writer' else windows_office_host._load_dependencies(COMPONENT[kind])
        assert deps.process_image(identity['office_pid']) == identity['office_executable'], 'Office process was terminated'
        try:
            api.create_document(kind, engine='msoffice')
        except Exception as exc:
            assert str(getattr(exc, 'code', '')).endswith('_QUARANTINED'), repr(exc)
        else:
            raise AssertionError('component quarantine did not block a new native session')
        write_json(Path(output) / 'fault-evidence.json', {'checks':{'exact_python_exit':True,
                   'office_alive':True, 'deadline_enforced':True, 'quarantine_blocks_new_session':True},
                   'recovery':recovery, 'quarantine_path':str(session._lock.quarantine_path)})
        return True
    finally:
        proxy._spawn_worker = original
        if session is not None:
            shutil.copytree(session.staging_root, Path(output) / 'fault-runtime')
        # Never close an uncertain native document or clear its quarantine.


def delayed_worker(job, plugin):
    import_api(plugin)
    from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker
    original = worker.default_factory
    def factory(*args, **kwargs):
        session = original(*args, **kwargs)
        inspect = session.inspect_document
        def stalled(*a, **k):
            time.sleep(60)
            return inspect(*a, **k)
        session.inspect_document = stalled
        return session
    worker.default_factory = factory
    # serve's default factory was bound at function definition; module rebinding
    # alone cannot inject this fault. Pass the exact factory to the protocol.
    job = Path(job).resolve(strict=True)
    tempfile.tempdir = str(job)
    for key in ('TMP', 'TEMP', 'TMPDIR'): os.environ[key] = str(job)
    worker.serve(sys.stdin.buffer, sys.stdout.buffer, job, factory=factory)


def native_case(plugin, output, kind, timeout, *, fault=False):
    output = Path(output)
    report = {'passed':False, 'kind':kind, 'checks':{}, 'fault':fault,
              'platform':sys.platform, 'engine':'msoffice', 'component':COMPONENT[kind],
              'capability_checks':{}, 'runner_sha256':digest(__file__),
              'baseline_scope':'registered application plus exact sentinel application; process image inventory',
              'office_termination_attempted':False}
    sentinels = None
    try:
        if sys.platform != 'win32': raise RuntimeError('Native acceptance requires Windows')
        api = import_api(plugin)
        report['source_digest'] = bound_source_digest(plugin)
        from skills.WPSComposer.scripts.msoffice.windows_office_runtime import _component_root
        write_json(output / 'recovery-root.json', {'root':str(_component_root(COMPONENT[kind]))})
        report['imported_api'] = str(Path(api.__file__).resolve())
        report['python'] = sys.executable
        sentinels = NativeSentinels(kind, output)
        sentinels.start()
        if fault:
            report['checks']['deadline_failure'] = fault_case(api, output, kind)
        else:
            report['checks'].update(standard_case(api, output, kind, time.monotonic() + timeout - 20))
        sentinels.verify(allow_owned_extra=fault)
        report['checks'].update(sentinels_preserved=True, baseline_preserved=True)
    except BaseException as exc:
        report['error'] = {'type':type(exc).__name__, 'message':str(exc)}
        (output / 'failure.txt').write_text(traceback.format_exc(), encoding='utf-8')
    finally:
        if sentinels is not None:
            try:
                report['checks']['sentinels_closed'] = sentinels.close(preserve_host=fault)
            except BaseException as exc:
                report['cleanup_error'] = {'type':type(exc).__name__, 'message':str(exc)}
        if sentinels is not None and report['checks'].get('sentinels_closed') is not True:
            try:
                from skills.WPSComposer.scripts.msoffice.windows_office_runtime import _component_root
                retain_quarantine(_component_root(COMPONENT[kind]), {
                    'reason':'Native acceptance sentinel cleanup was not verified',
                    'evidence_path':str(output), 'component':COMPONENT[kind]})
            except BaseException as exc:
                report['quarantine_error'] = str(exc)
        if report.get('source_digest'):
            try:
                report['source_digest_after'] = bound_source_digest(plugin)
                report['checks']['source_unchanged'] = report['source_digest_after'] == report['source_digest']
            except BaseException as exc:
                report['source_digest_error'] = str(exc)
                report['checks']['source_unchanged'] = False
        required = {'deadline_failure','sentinels_preserved','baseline_preserved','sentinels_closed','source_unchanged'} if fault else REQUIRED
        report['passed'] = not report.get('error') and not report.get('cleanup_error') and all(report['checks'].get(k) is True for k in required)
        # These exact format cases have direct artifact/reopen evidence. Do not
        # infer whole business-method, property, plan, or common-lifecycle rows.
        if report['passed'] and not fault:
            component, suffix = COMPONENT[kind], '.' + SUFFIX[kind]
            report['capability_checks'] = {
                f'format.{component}.inspect.{suffix}':['public_create_save_reopen','source_unchanged'],
                f'format.{component}.edit.{suffix}':['public_edit_reopen','source_preserved','source_unchanged'],
                f'format.{component}.save.{suffix}':['public_create_save_reopen','proxy_identity_close','source_unchanged'],
                f'format.{component}.generate.{suffix}':['public_generate','pdf_contents','source_unchanged'],
                f'format.{component}.convert_to_pdf.{suffix}':['public_convert_pdf','pdf_contents','source_unchanged'],
            }
        report['hashes'] = {str(p.relative_to(output)):digest(p) for p in output.rglob('*') if p.is_file()}
        write_json(output / 'case-report.json', report)
    return report


def run_probe(plugin, output, *, timeout=1800, kinds=KINDS, fault_timeout=False, execute=run_bounded):
    if type(timeout) not in (int, float) or not math.isfinite(timeout) or not 30 <= timeout <= 3600:
        raise ValueError('timeout must be finite, between 30 and 3600 seconds per app')
    plugin, output = Path(plugin).resolve(strict=True), Path(output).resolve()
    if output.exists(): raise FileExistsError(output)
    if not kinds or any(kind not in KINDS for kind in kinds) or len(set(kinds)) != len(kinds):
        raise ValueError('select unique writer, sheet, or slide kinds')
    output.mkdir(parents=True)
    shutil.copy2(__file__, output / 'runner-source.py')
    report = {'passed':False, 'plugin':str(plugin), 'runner_sha256':digest(__file__),
              'fault_timeout':'requested' if fault_timeout else 'not_requested', 'cases':{}}
    for fault in ([False, True] if fault_timeout else [False]):
        for kind in kinds:
            name = kind + ('-fault' if fault else '')
            case = output / name
            case.mkdir()
            command = [sys.executable, '-u', str(Path(__file__).resolve()), '--native-child', kind,
                       '--plugin', str(plugin), '--output', str(case), '--timeout', str(timeout)]
            if fault: command.append('--fault-timeout')
            outcome = execute(command, case, timeout=timeout)
            result = {'process':outcome, 'passed':False}
            try:
                native = json.loads((case / 'case-report.json').read_text(encoding='utf-8'))
                required = {'deadline_failure','sentinels_preserved','baseline_preserved','sentinels_closed','source_unchanged'} if fault else REQUIRED
                result['native'] = native
                result['passed'] = (outcome['status'] == 'completed' and outcome['returncode'] == 0 and
                    native.get('platform') == 'win32' and native.get('engine') == 'msoffice' and
                    native.get('component') == COMPONENT[kind] and
                    isinstance(native.get('source_digest'), str) and len(native['source_digest']) == 64 and
                    native.get('source_digest_after') == native['source_digest'] and
                    native.get('passed') is True and all(native.get('checks', {}).get(k) is True for k in required))
            except (OSError, ValueError) as exc:
                result['error'] = str(exc)
            report['cases'][name] = result
            write_json(output / 'report.json', report)
            if not result['passed']:
                # No later native request after uncertain state or a failed gate.
                return report
    report['passed'] = bool(report['cases']) and all(c['passed'] for c in report['cases'].values())
    write_json(output / 'report.json', report)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument('--plugin', type=Path)
    source.add_argument('--checkout', type=Path)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--kind', choices=KINDS, action='append')
    parser.add_argument('--timeout', type=float, default=1800)
    parser.add_argument('--fault-timeout', action='store_true', help='Opt-in: retains uncertain owned Office documents and quarantine')
    parser.add_argument('--native-child', choices=KINDS, help=argparse.SUPPRESS)
    parser.add_argument('--delayed-worker', type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args()
    plugin = args.plugin or args.checkout
    if args.delayed_worker:
        delayed_worker(args.delayed_worker, plugin)
        return 0
    if args.output is None: parser.error('--output is required')
    if args.native_child:
        report = native_case(plugin, args.output, args.native_child, args.timeout, fault=args.fault_timeout)
    else:
        report = run_probe(plugin, args.output, timeout=args.timeout, kinds=args.kind or KINDS, fault_timeout=args.fault_timeout)
    print(json.dumps({'passed':report['passed'], 'report':str(args.output / 'report.json')}, ensure_ascii=False))
    return 0 if report['passed'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
