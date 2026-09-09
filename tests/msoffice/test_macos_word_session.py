"""Native-session contracts with only the AppleEvent transport replaced."""
from __future__ import annotations

import importlib
import json
from pathlib import Path
import subprocess
import time

import pytest


def module():
    assert importlib.util.find_spec('skills.WPSComposer.scripts.msoffice.macos_word_session') is not None
    return importlib.import_module('skills.WPSComposer.scripts.msoffice.macos_word_session')


@pytest.fixture
def host(monkeypatch, tmp_path):
    m = module()
    monkeypatch.setattr(m.sys, 'platform', 'darwin')
    app = tmp_path / 'Word.app'
    app.mkdir()
    root = tmp_path / 'container'
    root.mkdir()
    monkeypatch.setattr(m, 'WORD_APP', app)
    monkeypatch.setattr(m, '_temporary_root', lambda: root)
    monkeypatch.setattr(m, 'validate_native_input', lambda *a, **k: None)
    source = tmp_path / 'source.docx'
    source.write_bytes(b'original source')
    scripts = []

    def run(command, **kwargs):
        script = Path(command[-1]).read_text()
        scripts.append(script)
        stage = Path(command[-1]).parent
        rows = [['binding', 901, str(next(stage.glob('document-*.docx')))]] if 'WPSC_BIND_OPEN' in script else [['ok']]
        if 'WPSC_BIND_ATTACH' in script:
            rows = [['binding', 902, '', 'Unsaved sentinel', False]]
        if 'counts.paragraphs' in script:
            rows = [['doc', 0, 'counts.paragraphs', 0]]
        return subprocess.CompletedProcess(command, 0, json.dumps(['WPSCOMPOSER_WORD_SESSION_OK', rows]), '')

    monkeypatch.setattr(m.subprocess, 'run', run)
    return m, source, scripts


def test_open_uses_private_source_and_exact_window_binding(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        stage = session.staging_root
        assert Path(session._bound_path).read_bytes() == source.read_bytes()
        session._execute(['set x to name of boundDoc'])
        assert f'document "{Path(session._bound_path).name}"' in scripts[-1]
        assert 'posix full name of boundDoc' in scripts[-1]
        assert str(source) not in scripts[0]
    assert source.read_bytes() == b'original source'
    assert not stage.exists()
    assert 'close boundDoc saving no' in scripts[-1]
    assert all('quit' not in s.lower() and 'clipboard' not in s.lower() for s in scripts)


def test_attached_close_never_closes_or_saves_user_document(host):
    m, _, scripts = host
    session = m.MacWordSession.attach_active()
    count = len(scripts)
    assert not session.supports_attached_save_copy()
    with pytest.raises(m.NativeWordCapabilityError):
        session.save_copy('/tmp/not-created.docx')
    session.close()
    assert len(scripts) == count


def test_timeout_quarantines_stage_and_blocks_later_jobs(host, monkeypatch):
    m, source, scripts = host
    session = m.MacWordSession.open_document(source)
    stage = session.staging_root
    def timeout(*a, **k):
        raise subprocess.TimeoutExpired('osascript', 1)
    monkeypatch.setattr(m.subprocess, 'run', timeout)
    with pytest.raises(m.NativeWordTimeoutError):
        session.inspect_document()
    with pytest.raises(m.NativeWordError):
        session.close()
    assert stage.is_dir()
    assert session.lock.quarantine_path.is_file()
    with pytest.raises(m.NativeWordError):
        m.MacWordSession.open_document(source)


@pytest.mark.parametrize('cancelled', [KeyboardInterrupt('cancelled'), SystemExit('cancelled')])
def test_cancelled_bound_transport_quarantines_without_context_exit_retry(host, monkeypatch, cancelled):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    stage = session.staging_root
    calls = []

    def cancel(*args, **kwargs):
        calls.append(args)
        raise cancelled

    monkeypatch.setattr(m.subprocess, 'run', cancel)
    with pytest.raises(type(cancelled), match='cancelled'):
        with session:
            session.inspect_document()

    assert len(calls) == 1
    assert stage.is_dir()
    assert session.lock.quarantine_path.is_file()
    logs = list(stage.glob('*.log'))
    assert any(type(cancelled).__name__ in log.read_text() for log in logs)


def test_cancelled_open_transport_retains_script_diagnostic_and_original_exception(host, monkeypatch):
    m, source, _ = host
    calls = []

    def cancel(*args, **kwargs):
        calls.append(args)
        raise KeyboardInterrupt('cancelled while opening')

    monkeypatch.setattr(m.subprocess, 'run', cancel)
    with pytest.raises(KeyboardInterrupt, match='cancelled while opening'):
        m.MacWordSession.open_document(source)

    stages = list(m._temporary_root().glob('wpscomposer-session-*'))
    assert len(calls) == 1
    assert len(stages) == 1
    assert list(stages[0].glob('*.applescript'))
    logs = list(stages[0].glob('*.log'))
    assert logs and 'KeyboardInterrupt' in logs[-1].read_text()
    assert (m._temporary_root() / 'wpscomposer-native-word.lock.quarantine').is_file()


def test_format_patch_uses_native_dictionary_without_selection(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        result = session.apply_format_patch('paragraph:2', font={'name': '仿宋', 'bold': True, 'color': '#123456'}, paragraph={'alignment': 3, 'first_line_indent': 24})
        assert result['rejected'] == []
        script = scripts[-1]
        assert 'text object of paragraph 2 of boundDoc' in script
        assert 'font object of targetRange' in script
        assert 'set east asian name' not in script
        assert 'align paragraph justify' in script
        assert 'selection' not in script
        assert 'set content' not in script


def test_unsupported_patch_fails_before_any_mutation(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        count = len(scripts)
        result = session.apply_format_patch('paragraph:1', text='must not change', font={'unknown': 7})
        assert result['accepted'] == []
        assert 'font.unknown' in result['rejected']
        assert len(scripts) == count


def test_read_only_session_rejects_mutation_before_transport(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source, read_only=True) as session:
        count = len(scripts)
        with pytest.raises(ValueError, match='read.only'):
            session.apply_format_patch('paragraph:1', text='changed')
        with pytest.raises(ValueError, match='read.only'):
            session.apply_structural_op({'op': 'remove', 'target': 'paragraph:1'})
        assert len(scripts) == count


def test_structure_clone_uses_formatted_text_not_clipboard(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        session.apply_structural_op({'op': 'clone', 'target': 'paragraph:1', 'to': 'end'})
        assert 'formatted text of' in scripts[-1]
        assert 'clipboard' not in scripts[-1].lower()
        assert 'paste' not in scripts[-1].lower()


def test_snapshot_decode_preserves_unicode_and_format_schema():
    m = module()
    result = m._decode_snapshot([
        ['doc', 0, 'name', 'test.docx'], ['doc', 0, 'saved', False],
        ['doc', 0, 'counts.paragraphs', 1],
        ['paragraphs', 1, 'text', '中文\t"test"\r'],
        ['paragraphs', 1, 'font.bold', True],
        ['paragraphs', 1, 'paragraph.alignment', 3],
    ])
    assert result['kind'] == 'writer'
    assert result['paragraphs'][0]['id'] == 'paragraph:1'
    assert result['paragraphs'][0]['text'] == '中文\t"test"'
    assert result['paragraphs'][0]['font']['bold'] is True


@pytest.mark.parametrize('op,expected', [
    ({'op': 'insert', 'type': 'table', 'props': {'rows': 2, 'cols': 2, 'data': [['A', 'B'], ['甲', '乙']]}}, 'make new table'),
    ({'op': 'insert', 'type': 'textbox', 'props': {'text': 'Box', 'width': 120, 'height': 40}}, 'make new text box'),
    ({'op': 'remove', 'target': 'table:1'}, 'delete targetObject'),
    ({'op': 'remove', 'target': 'shape:1'}, 'delete targetObject'),
])
def test_structural_objects_use_native_bound_document(host, op, expected):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        result = session.apply_structural_op(op)
        assert expected in scripts[-1]
        assert result['reinspect_required']


def test_section_columns_and_shape_wrap_are_not_silently_ignored(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        result = session.apply_format_patch('section:1', columns=2)
        assert result == {'accepted': ['columns'], 'rejected': []}
        assert 'set number of text columns targetSetup number of columns 2' in scripts[-1]
        result = session.apply_format_patch('shape:1', wrap=3)
        assert result == {'accepted': ['wrap'], 'rejected': []}
        assert 'wrap none' in scripts[-1]


def test_table_cell_snapshot_decodes_native_coordinates():
    m = module()
    result = m._decode_snapshot([
        ['tables', 1, 'rows', 2], ['tables', 1, 'columns', 2],
        ['cells', '1,2,1', 'text', 'Merged cell\r\x07'],
        ['cells', '1,2,1', 'fill.color', [65535, 0, 0]],
    ])
    assert result['tables'][0]['cells'][0] == {
        'id': 'table:1/cell:2,1', 'row': 2, 'column': 1,
        'text': 'Merged cell', 'fill': {'color': '#FF0000'},
    }


def test_save_copy_publishes_private_bytes_without_rebinding_or_touching_source(host, monkeypatch, tmp_path):
    m, source, scripts = host
    monkeypatch.setattr(m, 'validate_before_deadline', lambda *a: None)
    with m.MacWordSession.open_document(source) as session:
        Path(session._bound_path).write_bytes(b'edited bytes')
        output = tmp_path / 'copy.docx'
        assert session.save_copy(output) == str(output)
        assert output.read_bytes() == b'edited bytes'
        assert source.read_bytes() == b'original source'
        assert 'save as' not in scripts[-1]
        assert session._bound_path != str(output)


def test_rejected_source_does_not_acquire_lock_or_launch(host, monkeypatch):
    m, source, scripts = host
    def reject(*a, **k):
        raise ValueError('unsafe source')
    monkeypatch.setattr(m, 'validate_native_input', reject)
    with pytest.raises(ValueError):
        m.MacWordSession.open_document(source)
    assert scripts == []
    assert not (m._temporary_root() / 'wpscomposer-native-word.lock').exists()


def test_close_failure_retains_document_and_never_retries_unknown_host(host, monkeypatch):
    m, source, scripts = host
    session = m.MacWordSession.open_document(source)
    stage = session.staging_root
    monkeypatch.setattr(m.subprocess, 'run', lambda *a, **k: subprocess.CompletedProcess(a, 1, '', 'close failed'))
    with pytest.raises(m.NativeWordError):
        session.close()
    assert stage.is_dir()
    assert session.lock.quarantine_path.exists()
    assert session.lock.file is None


def test_missing_native_window_id_allows_only_private_owned_binding(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        session._bind([['binding', None, session._bound_path]])
        assert 'posix full name of boundDoc' in '\n'.join(session._binding())
    attached = m.MacWordSession()
    with pytest.raises(m.NativeWordCapabilityError):
        attached._bind([['binding', None, 'Unsaved1', 'Unsaved1']])


def test_failed_operation_retains_diagnostics_after_verified_document_close(host, monkeypatch):
    m, source, scripts = host
    session = m.MacWordSession.open_document(source)
    stage = session.staging_root
    original = m.subprocess.run
    monkeypatch.setattr(m.subprocess, 'run', lambda *a, **k: subprocess.CompletedProcess(a, 1, '', 'native property failed'))
    with pytest.raises(m.NativeWordError):
        session.apply_format_patch('paragraph:1', font={'bold': True})
    monkeypatch.setattr(m.subprocess, 'run', original)
    session.close()
    assert list(stage.glob('*.log'))
    assert not session.lock.quarantine_path.exists()


def test_save_preflight_rejects_unsupported_destination_before_transport(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        count = len(scripts)
        session.preflight_save()
        session.preflight_save('/tmp/output.docx')
        session.preflight_save(source, overwrite=True)
        with pytest.raises(FileExistsError):
            session.preflight_save(source, overwrite=False)
        with pytest.raises(m.NativeWordCapabilityError):
            session.preflight_save('/tmp/output.docm')
        assert len(scripts) == count
    attached = m.MacWordSession()
    with pytest.raises(m.NativeWordCapabilityError):
        attached.preflight_save()
    with pytest.raises(m.NativeWordCapabilityError):
        attached.preflight_save('/tmp/output.docx')


def test_image_insertion_uses_validated_private_copy(host, tmp_path):
    from PIL import Image
    m, source, scripts = host
    image = tmp_path / 'image.png'
    Image.new('RGB', (12, 12), 'red').save(image)
    with m.MacWordSession.open_document(source) as session:
        session.apply_structural_op({'op': 'insert', 'type': 'image', 'props': {'path': str(image)}})
        assert 'make new inline picture' in scripts[-1]
        assert 'save with document:true' in scripts[-1]
        assert str(image) not in scripts[-1]
        assert list(session.staging_root.glob('image-*.png'))


def test_paragraph_text_patch_preserves_paragraph_boundary(host):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        session.apply_format_patch('paragraph:2', text='replacement')
        script = scripts[-1]
        assert 'end (replacementEnd - 1)' in script
        assert 'set content of replacementRange to "replacement"' in script


def test_shape_inspection_includes_type_and_native_fill_line(host, monkeypatch):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        session.inspect_document()
        script = scripts[-1]
        assert 'shape type of targetObject' in script
        assert 'fore color of fill format of targetObject' in script
        assert 'weight of line format of targetObject' in script
        assert 'wrap type of wrap format of targetObject' in script
        assert 'set targetCell to cell cellIndex of text object of targetObject' in script


def test_prelaunch_copy_failure_does_not_quarantine_unstarted_word(host, monkeypatch):
    m, source, scripts = host
    def fail(*a, **k):
        raise OSError('disk failure')
    monkeypatch.setattr(m, 'copy_file_before_deadline', fail)
    with pytest.raises(OSError):
        m.MacWordSession.open_document(source)
    assert scripts == []
    assert not (m._temporary_root() / 'wpscomposer-native-word.lock.quarantine').exists()
    assert not list(m._temporary_root().glob('wpscomposer-session-*'))


def test_saved_paragraph_ids_are_emitted_only_with_verified_count(host, monkeypatch):
    m, source, scripts = host
    import skills.WPSComposer.scripts.writer as writer
    monkeypatch.setattr(writer, 'read_paraids_from_docx', lambda path: ['ABCDEF01'])
    monkeypatch.setattr(m, '_decode_snapshot', lambda rows: {'kind': 'writer', 'counts': {'paragraphs': 1}, 'paragraphs': [{'id': 'paragraph:1', 'index': 1}]})
    with m.MacWordSession.open_document(source) as session:
        assert session.inspect_document()['paragraphs'][0]['id'] == 'paragraph:@paraId=ABCDEF01'
        session._structural_changed = True
        assert session.inspect_document()['paragraphs'][0]['id'] == 'paragraph:1'


def test_session_deadline_is_shared_by_copy_validation_and_publication(host, monkeypatch, tmp_path):
    m, source, scripts = host
    deadlines = []
    copy = m.copy_file_before_deadline
    monkeypatch.setattr(m, 'copy_file_before_deadline', lambda *a, **k: (deadlines.append(k['deadline']), copy(*a, **k))[1])
    monkeypatch.setattr(m, 'validate_native_input', lambda *a, **k: deadlines.append(k['deadline']))
    monkeypatch.setattr(m, 'publish_artifact', lambda src, dst, **k: (deadlines.append(k['deadline']), copy(src, dst, deadline=k['deadline']), dst)[2])
    with m.MacWordSession.open_document(source) as session:
        deadline = session.publication_deadline
        session.save(tmp_path / 'output.docx')
        assert all(value == deadline for value in deadlines)
        assert deadline <= time.monotonic() + 600


def test_expired_session_cannot_refresh_budget_for_inspect_or_close(host, monkeypatch):
    m, source, scripts = host
    session = m.MacWordSession.open_document(source)
    session._deadline = time.monotonic() - 1
    count = len(scripts)
    with pytest.raises(m.NativeWordTimeoutError):
        session.inspect_document()
    assert len(scripts) == count
    with pytest.raises(m.NativeWordError):
        session.close()
    assert session.staging_root.exists()
    assert session.lock.quarantine_path.exists()


def test_private_document_basename_is_unique_across_sessions(host):
    m, source, _ = host
    with m.MacWordSession.open_document(source) as first:
        name = Path(first._bound_path).name
    with m.MacWordSession.open_document(source) as second:
        assert name != Path(second._bound_path).name
        assert name != 'document.docx'


@pytest.mark.parametrize('response', ['[]', '[["ok"]]', '["WPSCOMPOSER_WORD_SESSION_OK",null]'])
def test_missing_or_invalid_completion_envelope_quarantines(host, monkeypatch, response):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    monkeypatch.setattr(m.subprocess, 'run', lambda *a, **k: subprocess.CompletedProcess(a, 0, response, ''))
    with pytest.raises(m.NativeWordError):
        session.apply_format_patch('paragraph:1', text='change')
    assert session._quarantined
    assert session.staging_root.exists()
    with pytest.raises(m.NativeWordError):
        session.close()


def test_timeout_preserves_partial_native_output(host, monkeypatch):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    def timeout(*a, **k):
        raise subprocess.TimeoutExpired('osascript', 1, output=b'partial stdout', stderr=b'partial stderr')
    monkeypatch.setattr(m.subprocess, 'run', timeout)
    with pytest.raises(m.NativeWordTimeoutError):
        session.inspect_document()
    logs = '\n'.join(p.read_text() for p in session.staging_root.glob('*.log'))
    assert 'partial stdout' in logs and 'partial stderr' in logs
    with pytest.raises(m.NativeWordError):
        session.close()


def test_save_never_overwrites_existing_path_even_source(host, monkeypatch, tmp_path):
    m, source, scripts = host
    monkeypatch.setattr(m, 'validate_before_deadline', lambda *a: None)
    existing = tmp_path / 'existing.docx'
    existing.write_bytes(b'other document')
    with m.MacWordSession.open_document(source) as session:
        count = len(scripts)
        for path in (source, existing):
            with pytest.raises(FileExistsError):
                session.save(path)
        assert len(scripts) == count
        assert existing.read_bytes() == b'other document'


def test_save_current_requires_unchanged_original_source(host, monkeypatch):
    m, source, scripts = host
    monkeypatch.setattr(m, 'validate_before_deadline', lambda *a: None)
    with m.MacWordSession.open_document(source) as session:
        Path(session._bound_path).write_bytes(b'our edits')
        assert session.save_current() == str(source)
        assert source.read_bytes() == b'our edits'
        Path(session._bound_path).write_bytes(b'second edits')
        source.write_bytes(b'external edits')
        count = len(scripts)
        with pytest.raises(ValueError, match='changed'):
            session.save_current()
        assert len(scripts) == count
        assert source.read_bytes() == b'external edits'


def test_pdf_destination_preflight_runs_before_native_save(host, tmp_path):
    m, source, scripts = host
    existing = tmp_path / 'existing.pdf'
    existing.write_bytes(b'other PDF')
    with m.MacWordSession.open_document(source) as session:
        count = len(scripts)
        with pytest.raises(m.NativeWordCapabilityError):
            session.export_pdf(tmp_path / 'wrong.docx')
        with pytest.raises(FileExistsError):
            session.export_pdf(existing)
        assert len(scripts) == count
        assert existing.read_bytes() == b'other PDF'


def test_save_current_rechecks_source_after_native_save(host, monkeypatch):
    m, source, _ = host
    monkeypatch.setattr(m, 'validate_before_deadline', lambda *a: None)
    with m.MacWordSession.open_document(source) as session:
        execute = session._execute
        def external_edit(lines, **kwargs):
            result = execute(lines, **kwargs)
            if 'save boundDoc' in lines:
                source.write_bytes(b'external edit during save')
            return result
        monkeypatch.setattr(session, '_execute', external_edit)
        with pytest.raises(ValueError, match='changed'):
            session.save_current()
        assert source.read_bytes() == b'external edit during save'


def test_failed_publication_preserves_private_edits_after_close(host, monkeypatch, tmp_path):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    stage = session.staging_root
    private = Path(session._bound_path)
    private.write_bytes(b'recoverable edits')
    def failed_publish(*a, **k):
        raise OSError('destination unavailable')
    monkeypatch.setattr(m, 'publish_artifact', failed_publish)
    with pytest.raises(OSError):
        session.save(tmp_path / 'output.docx')
    session.close()
    assert stage.exists()
    assert private.read_bytes() == b'recoverable edits'
    assert not session.lock.quarantine_path.exists()


@pytest.mark.skipif(not Path('/Applications/Microsoft Word.app/Contents/Resources/Word.sdef').is_file(), reason='Mac Word dictionary required for compile-only test')
def test_native_dictionary_compiles_range_snapshot_and_cell_patch(host, tmp_path):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        session._execute(session._range('paragraph:1') + session._range_rows('paragraphs', 1))
        snapshot_source = scripts[-1].replace(str(m.WORD_APP), '/Applications/Microsoft Word.app')
        session.apply_format_patch('table:1/cell:1,1', vertical_alignment=1, paragraph={'line_spacing_rule': 1})
        sources = [snapshot_source, scripts[-1].replace(str(m.WORD_APP), '/Applications/Microsoft Word.app')]
        session.inspect_document()
        sources.append(scripts[-1].replace(str(m.WORD_APP), '/Applications/Microsoft Word.app'))
        for i, script in enumerate(sources):
            path = tmp_path / f'compile-{i}.applescript'
            path.write_text(script)
            process = subprocess.Popen(['/usr/bin/osacompile', '-o', str(path.with_suffix('.scpt')), str(path)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            out, error = process.communicate(timeout=15)
            assert process.returncode == 0, error


def test_new_word_uses_native_make_and_exact_private_binding(host, monkeypatch):
    m, _, scripts = host
    def execute(self, lines, **kwargs):
        scripts.append('\n'.join(lines))
        if not kwargs.get('bind', True):
            self._private_path.write_bytes(b'native created')
            return [['binding', None, str(self._private_path)]]
        return [['ok']]
    monkeypatch.setattr(m.MacWordSession, '_execute', execute)
    with m.MacWordSession.new_document() as session:
        assert session._owns_doc
        assert session._source_path is None
        assert 'make new document' in scripts[0]
        assert 'save as boundDoc' in scripts[0]
        assert str(session._private_path) in scripts[0]
        with pytest.raises(ValueError, match='explicit'):
            session.save_current()
    assert 'close boundDoc saving no' in scripts[-1]


@pytest.mark.parametrize('kind', ['paragraph','heading'])
def test_insert_paragraph_establishes_native_boundary_before_text(host, kind):
    m, source, scripts = host
    with m.MacWordSession.open_document(source) as session:
        session.apply_structural_op({'op':'insert','type':kind,'props':{'text':'Separate'}})
        script = scripts[-1]
        assert 'set precedingRange to create range boundDoc' in script
        assert 'set insertionPoint to insertionPoint + 1' in script
        assert script.index('set insertionPoint to insertionPoint + 1') < script.index('"Separate"')


@pytest.mark.parametrize('group,values', [('font',{'size':-1}),('font',{'size':0}),('fill',{'transparency':1.1}),('geometry',{'width':-1}),('line',{'weight':-2})])
def test_invalid_numeric_patch_rejects_whole_mutation(host,group,values):
    m,source,scripts=host
    with m.MacWordSession.open_document(source) as session:
        count=len(scripts)
        result=session.apply_format_patch('shape:1',text='must remain unchanged',**{group:values})
        assert result['accepted']==[] and result['rejected']
        assert len(scripts)==count


def test_close_save_conflict_keeps_session_lock_until_discard(host):
    m,source,scripts=host
    session=m.MacWordSession.open_document(source)
    source.write_bytes(b'external update')
    with pytest.raises(ValueError,match='changed'):session.close(save_changes=True)
    assert not session._closed
    assert session.lock.file is not None
    session.close()
    assert session._closed


def test_attached_binding_requires_native_read_only_state_and_checks_source_hash(host,monkeypatch):
    m,source,scripts=host
    session=m.MacWordSession();session._prepare()
    try:
        session._bind([['binding',901,str(source),source.name,False]])
        assert session._source_digest==session._digest(source)
        source.write_bytes(b'external update')
        count=len(scripts)
        with pytest.raises(ValueError,match='changed'):session.preflight_save()
        with pytest.raises(ValueError,match='changed'):session.save_current()
        assert len(scripts)==count
    finally:session.close()


def test_attached_native_read_only_blocks_mutation(host):
    m,source,scripts=host
    session=m.MacWordSession();session._prepare()
    try:
        session._bind([['binding',901,str(source),source.name,True]])
        count=len(scripts)
        with pytest.raises(ValueError,match='read-only'):session.apply_format_patch('paragraph:1',text='blocked')
        with pytest.raises(ValueError,match='read-only'):session.preflight_save()
        assert len(scripts)==count
    finally:session.close()


@pytest.mark.parametrize('target,text',[('paragraph:1','two\rparagraphs'),('range:0-9','plain replacement'),('selection','plain replacement'),('table:1/cell:1,1','plain replacement')])
def test_text_boundary_changes_invalidate_saved_para_ids(host,target,text):
    m,source,scripts=host
    with m.MacWordSession.open_document(source) as session:
        session.apply_format_patch(target,text=text)
        assert session._structural_changed
        with pytest.raises(ValueError,match='stale'):session._range('paragraph:@paraId=00000001')


def test_attached_missing_read_only_state_rejected_before_mutation():
    m=module();session=m.MacWordSession()
    with pytest.raises(m.NativeWordCapabilityError):session._bind([['binding',901,'','Unsaved']])


def test_attached_save_refreshes_digest_and_checks_read_only(host,monkeypatch):
    m,source,scripts=host
    session=m.MacWordSession();session._prepare()
    try:
        session._bind([['binding',901,str(source),source.name,False]])
        execute=session._execute
        def saving(lines,**kwargs):
            result=execute(lines,**kwargs)
            if 'save boundDoc' in lines:source.write_bytes(b'our saved edit')
            return result
        monkeypatch.setattr(session,'_execute',saving)
        assert session.save_current()==str(source)
        assert session._source_digest==session._digest(source)
        assert 'if read only of boundDoc then error "WPSC_READ_ONLY"' in scripts[-1]
        session.preflight_save()
    finally:session.close()


def test_plain_single_paragraph_replacement_keeps_stable_ids(host):
    m,source,scripts=host
    with m.MacWordSession.open_document(source) as session:
        session.apply_format_patch('paragraph:1',text='Same paragraph')
        assert not session._structural_changed


@pytest.mark.parametrize('native', ['success', 'ordinary', 'native-timeout', 'timeout', 'cancel'])
@pytest.mark.parametrize('lost', ['log', 'quarantine', 'both'])
def test_submitted_diagnostic_failure_retains_original_outcome_and_blocks_reentry(host, monkeypatch, native, lost):
    """Disk failures must not hide the primary native error or allow another run."""
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    session._observed_field_topology = (('before', 1),)
    cancelled = KeyboardInterrupt('original cancellation')
    calls = []
    original_write = Path.write_text

    def write(path, value, **kwargs):
        if lost in ('log', 'both') and path.suffix == '.log':
            raise OSError('injected diagnostic disk full')
        return original_write(path, value, **kwargs)

    def quarantine(detail):
        raise OSError('injected quarantine disk full')

    def run(command, **kwargs):
        calls.append(command)
        if native == 'cancel':
            raise cancelled
        if native == 'timeout':
            raise subprocess.TimeoutExpired(command, 1, output=b'partial output')
        if native in ('ordinary', 'native-timeout'):
            code = -2700 if native == 'ordinary' else -1712
            return subprocess.CompletedProcess(command, 1, '', f'execution error: native ({code})')
        return subprocess.CompletedProcess(command, 0, json.dumps([m._COMPLETION_MARKER, [['ok']]]), '')

    monkeypatch.setattr(Path, 'write_text', write)
    monkeypatch.setattr(m.subprocess, 'run', run)
    if lost in ('quarantine', 'both'):
        monkeypatch.setattr(session.lock, 'quarantine', quarantine)
    # A normal result does not need quarantine unless its diagnostic fails.
    if native == 'success' and lost == 'quarantine':
        assert session._execute(['set nativeRows to {{"ok"}}']) == [['ok']]
        assert not session._quarantined
        return
    error_type = KeyboardInterrupt if native == 'cancel' else m.NativeWordTimeoutError if native in ('timeout', 'native-timeout') else m.NativeWordError
    with pytest.raises(error_type) as observed:
        session._execute(['set nativeRows to {{"ok"}}'])
    if native == 'cancel':
        assert observed.value is cancelled
    else:
        if native == 'ordinary':
            assert observed.value.code == 'NATIVE_WORD_EXECUTION_FAILED'
        if lost in ('log', 'both'):
            assert observed.value.diagnostic_path is None
        else:
            assert Path(observed.value.diagnostic_path).is_file()
        if lost in ('quarantine', 'both'):
            assert observed.value.quarantine_path is None
    if native == 'ordinary' and lost == 'quarantine':
        # A fully recorded ordinary error retains its existing nonquarantine
        # classification; this path never attempts the injected lock write.
        assert not session._quarantined and session._retain_evidence
        return
    assert session._quarantined and session._retain_evidence
    assert session._observed_field_topology == (('before', 1),)
    assert len(calls) == 1
    with pytest.raises(m.NativeWordError) as blocked:
        session._execute(['set nativeRows to {{"must not launch"}}'])
    assert blocked.value.code == 'NATIVE_WORD_QUARANTINED'
    assert len(calls) == 1


@pytest.mark.parametrize('failure', ['script', 'chmod', 'deadline'])
def test_transport_diagnostic_fix_does_not_turn_prelaunch_failure_into_submission(host, monkeypatch, failure):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    session._observed_field_topology = (('before', 1),)
    calls = []
    original_write = Path.write_text

    def write(path, value, **kwargs):
        if failure == 'script' and path.suffix == '.applescript':
            raise OSError('script disk full')
        return original_write(path, value, **kwargs)

    def chmod(path, mode):
        raise OSError('chmod failed')

    monkeypatch.setattr(Path, 'write_text', write)
    monkeypatch.setattr(m.subprocess, 'run', lambda *a, **k: calls.append(a))
    if failure == 'chmod':
        monkeypatch.setattr(m.os, 'chmod', chmod)
    if failure == 'deadline':
        session._deadline = -1
    with pytest.raises((OSError, m.NativeWordTimeoutError)):
        session._execute_topology_mutation(['set nativeRows to {{"ok"}}'])
    assert not calls and session._observed_field_topology == (('before', 1),)
    assert session._quarantined is (failure == 'deadline')
    assert not list(session.staging_root.glob('*.log'))[1:]


def test_partial_quarantine_record_is_not_advertised_by_later_retention(host, monkeypatch):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    path = session.lock.quarantine_path
    def partial_record(detail):
        path.write_text('{', encoding='utf-8')
        raise OSError('partial quarantine write')
    monkeypatch.setattr(session.lock, 'quarantine', partial_record)
    session._retain('first failure')
    session._retain('outer wrapper retains again')
    assert session._quarantined and path.is_file()
    assert session._quarantine_location() is None


@pytest.mark.parametrize('primary', ['timeout', 'cancel'])
def test_secondary_evidence_cancellation_cannot_replace_original_native_error(host, monkeypatch, primary):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    original = KeyboardInterrupt('primary cancellation')
    def run(*args, **kwargs):
        if primary == 'cancel':
            raise original
        raise subprocess.TimeoutExpired('osascript', 1)
    def cancelled_evidence(*args, **kwargs):
        raise SystemExit('secondary diagnostic cancellation')
    original_write = Path.write_text
    def write(path, data, **kwargs):
        if path.suffix == '.log':
            cancelled_evidence()
        return original_write(path, data, **kwargs)
    monkeypatch.setattr(m.subprocess, 'run', run)
    monkeypatch.setattr(Path, 'write_text', write)
    monkeypatch.setattr(session.lock, 'quarantine', cancelled_evidence)
    with pytest.raises(KeyboardInterrupt if primary == 'cancel' else m.NativeWordTimeoutError) as observed:
        session._execute(['set nativeRows to {{"ok"}}'])
    if primary == 'cancel':
        assert observed.value is original
    else:
        assert observed.value.diagnostic_path is observed.value.quarantine_path is None
    assert session._quarantined


def test_post_submission_log_failure_invalidates_content_observations(host, monkeypatch):
    m, source, _ = host
    session = m.MacWordSession.open_document(source)
    session._observed_field_topology = (('before', 1),)
    original_write = Path.write_text
    def write(path, data, **kwargs):
        if path.suffix == '.log':
            raise OSError('diagnostic failure')
        return original_write(path, data, **kwargs)
    monkeypatch.setattr(Path, 'write_text', write)
    with pytest.raises(m.NativeWordError):
        session._execute_topology_mutation(['set nativeRows to {{"ok"}}'])
    assert session._quarantined and not hasattr(session, '_observed_field_topology')
    assert not session._field_topology_mutation_pending
