"""Bound Microsoft sessions exercised against native COM boundary doubles."""
from pathlib import Path
from types import SimpleNamespace
import importlib
import json
import sys
import zipfile
import time

import pytest


def package(path, kind):
    root, main, content_type = {
        'writer': ('word', 'document', 'wordprocessingml.document'),
        'sheet': ('xl', 'workbook', 'spreadsheetml.sheet'),
        'slide': ('ppt', 'presentation', 'presentationml.presentation'),
    }[kind]
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + root + '/' + main + '.xml" ContentType="application/vnd.openxmlformats-officedocument.' + content_type + '.main+xml"/></Types>')
        archive.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="' + root + '/' + main + '.xml"/></Relationships>')
        archive.writestr(root + '/' + main + '.xml', '<' + main + '/>')
    return path


class Collection:
    def __init__(self, values=()):
        self.values = list(values)
    @property
    def Count(self):
        return len(self.values)
    def Item(self, i):
        return self.values[i - 1]
    __call__ = Item


class Doc:
    def __init__(self, app, kind, path):
        self.Application, self.kind = app, kind
        self.FullName, self.Name = str(path), Path(path).name
        self.Path = str(Path(path).parent) if Path(path).is_absolute() else ''
        self.Saved = False
        self.ReadOnly = False
        self.HasVBProject = False
        self.Excel4MacroSheets = self.Excel4IntlMacroSheets = Collection()
        self.token = object()
        self.closed = self.save_calls = 0
        self.Windows = Collection([SimpleNamespace(Hwnd=44, Presentation=self)])
        self.Windows.Item(1).Selection = SimpleNamespace(Document=self, Range=SimpleNamespace(Text='owned selection', Start=0, End=15))
        self.Worksheets = Collection()
        self.Paragraphs = Collection()
        self.Tables = self.Shapes = self.Sections = self.Slides = Collection()
        self.PageSetup = SimpleNamespace()
        self.data = 'initial'

    def Save(self):
        self.save_calls += 1
        self.Saved = True

    def SaveCopyAs(self, path, *args, **kwargs):
        package(Path(path), self.kind)

    def Close(self, **kwargs):
        self.closed += 1
        self.Application.documents.values.remove(self)


class Composer:
    def __init__(self, app, doc, kind):
        self._app, self._doc, self.kind = app, doc, kind
        self._owns_app = True
        self._owns_doc = True
        self._closed = False
        self.cleanup_error = None
        self._read_only = False
        self._com_initialized = True
        self._deps = SimpleNamespace(identity=lambda obj: obj.token, uninitialize=lambda: None)
        self._application_token = app.token
        self._document_token = doc.token
        self._baseline = []
        self.executable = app.Path + '\\' + {'writer': 'WINWORD.EXE', 'sheet': 'EXCEL.EXE', 'slide': 'POWERPNT.EXE'}[kind]
        self.pid = 22
        self.identity = SimpleNamespace(pid=22, executable=self.executable)
        self.fail_verify = False
    def _verify_document(self, *args):
        if self.fail_verify:
            raise RuntimeError('identity lost')
    def open_owned_document(self, path, read_only=True):
        self.opened_path = Path(path)
        self._doc.FullName = str(path)
        self._read_only = read_only
        self._doc.ReadOnly = read_only
    def save_docx(self, path):
        return self.save_native(path)
    save_xlsx = save_pptx = save_docx
    def save_native(self, path):
        package(Path(path), self.kind)
        self._doc.FullName = str(path)
        return str(path)
    def export_pdf(self, path):
        from pypdf import PdfWriter
        writer = PdfWriter()
        writer.add_blank_page(200, 200)
        with Path(path).open('wb') as stream:
            writer.write(stream)
        return str(path)
    def inspect_document(self, **kwargs):
        return {'kind': self.kind, 'path': self._doc.FullName, 'text': self._doc.data}
    def apply_format_patch(self, target, **kwargs):
        self._doc.data = kwargs.get('text', kwargs.get('value', 'formatted'))
        return {'accepted': list(kwargs), 'rejected': []}
    def apply_structural_op(self, op):
        self._doc.data = str(self._doc.data) + '|inserted'
        return {'type': op['type'], 'path': 'element:2'}
    def close(self, save_changes=False):
        assert not save_changes
        self._doc.Close()
        self._app.quit = self._owns_app
        self._closed = True


@pytest.fixture
def api():
    return importlib.import_module('skills.WPSComposer.scripts.msoffice.windows_document_api')


def app_doc(kind, path='unsaved sentinel'):
    app = SimpleNamespace(Name={'writer': 'Microsoft Word', 'sheet': 'Microsoft Excel', 'slide': 'Microsoft PowerPoint'}[kind],
        Path=r'C:\Office', token=object(), Hwnd=43, AutomationSecurity=1, quit=False)
    doc = Doc(app, kind, path)
    app.documents = Collection([doc])
    app.Documents = app.Workbooks = app.Presentations = app.documents
    app.ActiveDocument = app.ActiveWorkbook = app.ActivePresentation = doc
    app.ActiveWindow = doc.Windows.Item(1)
    app.Selection = SimpleNamespace(Parent=SimpleNamespace(Parent=doc))
    return app, doc


def cls(api, kind):
    return {'writer': api.WindowsWordSession, 'sheet': api.WindowsExcelSession,
            'slide': api.WindowsPowerPointSession}[kind]


def owned(api, monkeypatch, tmp_path, kind, read_only=False):
    suffix = {'writer': '.docx', 'sheet': '.xlsx', 'slide': '.pptx'}[kind]
    source = package(tmp_path / ('source' + suffix), kind)
    app, doc = app_doc(kind)
    composer = Composer(app, doc, kind)
    monkeypatch.setattr(api, '_create_owned_composer', lambda component, root: composer)
    return cls(api, kind).open_document(source, read_only=read_only), composer, source


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_owned_open_and_semantic_edit_use_private_copy_until_explicit_save(api, monkeypatch, tmp_path, kind):
    session, composer, source = owned(api, monkeypatch, tmp_path, kind)
    original = source.read_bytes()
    assert composer.opened_path != source and composer.opened_path.read_bytes() == original
    with session as same:
        assert same is session
        session.apply_format_patch('selection', **({'value': 42} if kind == 'sheet' else {'text': 'changed'}))
        session.apply_structural_op({'op': 'insert', 'type': 'paragraph' if kind == 'writer' else 'sheet' if kind == 'sheet' else 'slide'})
        assert 'inserted' in str(session.inspect_document()['text'])
        assert source.read_bytes() == original
    assert composer._closed and not composer.opened_path.exists()
    assert source.read_bytes() == original


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_read_only_sessions_reject_mutation_before_native_call(api, monkeypatch, tmp_path, kind):
    session, composer, source = owned(api, monkeypatch, tmp_path, kind, read_only=True)
    with pytest.raises(PermissionError):
        session.apply_format_patch('anything', text='changed')
    with pytest.raises(PermissionError):
        session.apply_structural_op({'op': 'remove', 'target': 'anything'})
    assert composer._doc.data == 'initial'
    session.close()


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_save_to_output_keeps_native_bound_to_private_file_and_source_unchanged(api, monkeypatch, tmp_path, kind):
    session, composer, source = owned(api, monkeypatch, tmp_path, kind)
    original = source.read_bytes()
    output = tmp_path / ('saved' + source.suffix)
    assert session.save(output) == str(output)
    assert output.exists() and Path(composer._doc.FullName) != output
    assert source.read_bytes() == original
    session.close()


def test_save_current_rejects_concurrent_source_change(api, monkeypatch, tmp_path):
    session, composer, source = owned(api, monkeypatch, tmp_path, 'sheet')
    source.write_bytes(b'new user edit')
    with pytest.raises(RuntimeError, match='changed'):
        session.save_current()
    assert source.read_bytes() == b'new user edit'
    session.close()


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_new_document_uses_one_owned_host_and_requires_explicit_first_save(api, monkeypatch, tmp_path, kind):
    app, doc = app_doc(kind)
    composer = Composer(app, doc, kind)
    calls=[]
    def create(component, root):
        calls.append((component, root));return composer
    monkeypatch.setattr(api, '_create_owned_composer', create)
    with cls(api, kind).new_document(visible=True) as session:
        assert len(calls)==1
        assert not hasattr(composer, 'opened_path')
        assert Path(doc.FullName).is_file()
        assert session.staging_root in Path(doc.FullName).parents
        assert session.inspect_document()['path'] is None
        with pytest.raises(ValueError, match='explicit'):
            session.save_current()
        with pytest.raises(ValueError, match='explicit'):
            session.preflight_save()
        assert app.Visible is True
        output=tmp_path/('new.'+{'writer':'docx','sheet':'xlsx','slide':'pptx'}[kind])
        assert session.save(output)==str(output)
        assert output.is_file()
        assert session.inspect_document()['path']==str(output)
    assert doc.closed==1
    assert not session.staging_root.exists()


def test_new_document_private_save_failure_retains_evidence_and_closes_owned_only(api, monkeypatch):
    app, doc=app_doc('writer')
    composer=Composer(app,doc,'writer')
    roots=[]
    monkeypatch.setattr(api,'_create_owned_composer',lambda component,root:roots.append(root) or composer)
    def fail(path):raise OSError('private save failed')
    monkeypatch.setattr(composer,'save_docx',fail)
    with pytest.raises(OSError):api.WindowsWordSession.new_document()
    assert doc.closed==1
    assert list(roots[0].glob('failure-*.json'))


@pytest.mark.parametrize('kind', ['writer','sheet','slide'])
def test_direct_save_copy_and_pdf_never_overwrite_existing_destination(api, monkeypatch, tmp_path, kind):
    session,composer,source=owned(api,monkeypatch,tmp_path,kind)
    original=source.read_bytes();bound=composer._doc.FullName
    pdf=tmp_path/'existing.pdf';pdf.write_bytes(b'keep PDF')
    for method,target in [(session.save,source),(session.save_copy,source),(session.export_pdf,pdf)]:
        with pytest.raises(FileExistsError):method(target)
    assert composer._doc.FullName==bound
    assert source.read_bytes()==original and pdf.read_bytes()==b'keep PDF'
    session.close()


def test_saved_output_conflict_prevents_explicit_current_replace(api, monkeypatch, tmp_path):
    session,composer,source=owned(api,monkeypatch,tmp_path,'writer')
    output=tmp_path/'saved.docx'
    session.save(output)
    output.write_bytes(b'external edit')
    with pytest.raises(RuntimeError,match='changed'):session.save_current()
    assert output.read_bytes()==b'external edit'
    session.close()


def test_windows_session_copy_hash_validation_and_publish_share_one_deadline(api, monkeypatch, tmp_path):
    observed=[]
    native_validate=api.validate_native_input
    monkeypatch.setattr(api,'validate_native_input',lambda *a,**k:(observed.append(k['deadline']),native_validate(*a,**k))[1])
    copy=api.copy_file_before_deadline
    monkeypatch.setattr(api,'copy_file_before_deadline',lambda *a,**k:(observed.append(k['deadline']),copy(*a,**k))[1])
    session,composer,source=owned(api,monkeypatch,tmp_path,'writer')
    deadline=session.publication_deadline
    publish=api.publish_artifact
    monkeypatch.setattr(api,'publish_artifact',lambda *a,**k:(observed.append(k['deadline']),publish(*a,**k))[1])
    session.save(tmp_path/'copy.docx')
    assert observed and all(value==deadline for value in observed)
    session._deadline=time.monotonic()-1
    with pytest.raises(TimeoutError):session.inspect_document()
    session._deadline=deadline
    session.close()


def test_current_save_detects_source_edit_during_native_save(api, monkeypatch, tmp_path):
    session,composer,source=owned(api,monkeypatch,tmp_path,'writer')
    original=composer.save_docx
    def save(path):
        original(path)
        source.write_bytes(b'user edit during native save')
    monkeypatch.setattr(composer,'save_docx',save)
    with pytest.raises(RuntimeError,match='changed'):session.save_current()
    assert source.read_bytes()==b'user edit during native save'
    stage=session.staging_root
    session.close()
    assert stage.exists() and list(stage.glob('output-*.docx'))


def attach(api, monkeypatch, kind, *, wrong_image=False):
    app, doc = app_doc(kind)
    executable = app.Path + '\\' + ('wps.exe' if wrong_image else {'writer': 'WINWORD.EXE', 'sheet': 'EXCEL.EXE', 'slide': 'POWERPNT.EXE'}[kind])
    deps = SimpleNamespace(initialize=lambda: None, uninitialize=lambda: None,
        active=lambda: app, identity=lambda obj: obj.token,
        process_image=lambda pid: executable, window_pid=lambda hwnd: 22,
        powerpoint_windows=lambda: [(44, SimpleNamespace(Application=app, Presentation=doc))])
    monkeypatch.setattr(api, '_attach_dependencies', lambda component: deps)
    return app, doc


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_attach_close_never_saves_closes_or_quits_live_document(api, monkeypatch, kind):
    app, doc = attach(api, monkeypatch, kind)
    session = cls(api, kind).attach_active()
    assert session._owns_doc is False
    session.close(save_changes=False)
    assert not doc.closed and not doc.save_calls and not doc.Saved and not app.quit
    assert app.AutomationSecurity == 1


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_wrong_native_executable_prevents_attachment(api, monkeypatch, kind):
    app, doc = attach(api, monkeypatch, kind, wrong_image=True)
    with pytest.raises(api.DocumentIdentityError):
        cls(api, kind).attach_active()
    assert not doc.closed and not app.quit


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_active_document_switch_cannot_redirect_selection_edit(api, monkeypatch, kind):
    app, doc = attach(api, monkeypatch, kind)
    session = cls(api, kind).attach_active()
    other = Doc(app, kind, 'other user document')
    app.documents.values.append(other)
    app.ActiveDocument = app.ActiveWorkbook = app.ActivePresentation = other
    app.ActiveWindow = other.Windows.Item(1)
    app.Selection = SimpleNamespace(Parent=SimpleNamespace(Parent=other))
    with pytest.raises(api.DocumentIdentityError):
        session.apply_format_patch('selection', text='must not redirect')
    assert doc.data == other.data == 'initial'
    session.close()
    assert not other.closed and not doc.closed


@pytest.mark.parametrize('kind', ['sheet', 'slide'])
def test_attached_copy_preserves_unsaved_binding(api, monkeypatch, tmp_path, kind):
    app, doc = attach(api, monkeypatch, kind)
    session = cls(api, kind).attach_active()
    output = tmp_path / ('copy.xlsx' if kind == 'sheet' else 'copy.pptx')
    original = (doc.FullName, doc.Saved)
    assert session.supports_attached_save_copy()
    assert session.save_copy(output) == str(output)
    assert (doc.FullName, doc.Saved) == original
    assert output.exists()
    session.close()


def test_word_attached_copy_rejects_without_saveas_rebind(api, monkeypatch, tmp_path):
    app, doc = attach(api, monkeypatch, 'writer')
    session = api.WindowsWordSession.attach_active()
    assert not session.supports_attached_save_copy()
    with pytest.raises(RuntimeError):
        session.save_copy(tmp_path / 'copy.docx')
    assert doc.FullName == 'unsaved sentinel' and not doc.Saved
    session.close()


def test_stale_identity_refuses_write_and_retains_private_recovery(api, monkeypatch, tmp_path):
    session, composer, source = owned(api, monkeypatch, tmp_path, 'sheet')
    composer.fail_verify = True
    with pytest.raises(RuntimeError):
        session.apply_format_patch('sheet:1', name='changed')
    assert composer._doc.data == 'initial'
    assert session.staging_root.exists()


def test_real_excel_format_patch_updates_only_addressed_owned_cell(api, monkeypatch):
    app, doc = attach(api, monkeypatch, 'sheet')
    cell = SimpleNamespace(Value=1, Formula='', Font=SimpleNamespace(Bold=False))
    foreign = SimpleNamespace(Value=90, Formula='', Font=SimpleNamespace(Bold=False))
    worksheet = SimpleNamespace(Range=lambda address: cell if address == 'B2' else foreign)
    doc.Worksheets.values.append(worksheet)
    session = api.WindowsExcelSession.attach_active()
    result = session.apply_format_patch('sheet:1/cell:B2', value=7, font={'bold': True})
    assert not result['rejected'] and cell.Value == 7 and cell.Font.Bold
    assert foreign.Value == 90 and not foreign.Font.Bold
    session.close()


def test_real_powerpoint_structural_delete_removes_only_bound_slide(api, monkeypatch):
    app, doc = attach(api, monkeypatch, 'slide')
    slide = SimpleNamespace()
    slide.Delete = lambda: doc.Slides.values.remove(slide)
    doc.Slides.values.append(slide)
    session = api.WindowsPowerPointSession.attach_active()
    result = session.apply_structural_op({'op': 'remove', 'target': 'slide:1'})
    assert result['removed'] == 'slide:1' and doc.Slides.Count == 0
    assert not doc.closed and not app.quit
    session.close()


def test_real_word_text_edit_uses_bound_range(api, monkeypatch):
    app, doc = attach(api, monkeypatch, 'writer')
    rng = SimpleNamespace(Text='before', Font=SimpleNamespace(), ParagraphFormat=SimpleNamespace())
    doc.Paragraphs.values.append(SimpleNamespace(Range=rng))
    session = api.WindowsWordSession.attach_active()
    result = session.apply_format_patch('paragraph:1', text='after')
    assert rng.Text == 'after' and 'text' in result['accepted']
    session.close()


def test_changed_savecopy_binding_never_publishes_result(api, monkeypatch, tmp_path):
    app, doc = attach(api, monkeypatch, 'sheet')
    session = api.WindowsExcelSession.attach_active()
    def bad_copy(path):
        package(Path(path), 'sheet')
        doc.FullName = path
    doc.SaveCopyAs = bad_copy
    target = tmp_path / 'must-not-publish.xlsx'
    with pytest.raises(api.DocumentIdentityError):
        session.save_copy(target)
    assert not target.exists() and session.staging_root.exists()
    session.close()
    assert not doc.closed


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_live_macro_project_blocks_edits_without_changing_shared_security(api, monkeypatch, kind):
    app, doc = attach(api, monkeypatch, kind)
    session = cls(api, kind).attach_active()
    doc.HasVBProject = True
    with pytest.raises(PermissionError, match='macro'):
        session.apply_format_patch('selection', text='blocked')
    assert doc.data == 'initial' and app.AutomationSecurity == 1
    session.close()
    assert not doc.closed


@pytest.mark.parametrize('field', ['Excel4MacroSheets', 'Excel4IntlMacroSheets'])
def test_live_excel_macro_sheets_block_savecopy_before_native_call(api, monkeypatch, tmp_path, field):
    app, doc = attach(api, monkeypatch, 'sheet')
    session = api.WindowsExcelSession.attach_active()
    setattr(doc, field, Collection([object()]))
    doc.SaveCopyAs = lambda *args: pytest.fail('macro workbook must not reach native save')
    with pytest.raises(PermissionError, match='macro'):
        session.save_copy(tmp_path / 'copy.xlsx')
    session.close()


def test_unverifiable_live_macro_status_fails_before_native_mutation(api, monkeypatch):
    app, doc = attach(api, monkeypatch, 'sheet')
    session = api.WindowsExcelSession.attach_active()
    del doc.HasVBProject
    with pytest.raises(PermissionError, match='macro'):
        session.apply_format_patch('selection', value=3)
    session.close()


def test_native_output_with_hidden_macro_part_never_publishes(api, monkeypatch, tmp_path):
    session, composer, source = owned(api, monkeypatch, tmp_path, 'sheet')
    original_save = composer.save_native
    def unsafe_save(path):
        original_save(path)
        with zipfile.ZipFile(path, 'a') as archive:
            archive.writestr('xl/vbaProject.bin', b'unsafe')
    composer.save_native = unsafe_save
    target = tmp_path / 'output.xlsx'
    with pytest.raises(Exception):
        session.save(target)
    assert not target.exists() and session.staging_root.exists()
    session.close()


def test_word_security_setup_failure_closes_new_owned_document(api, monkeypatch, tmp_path):
    class DeniedSecurity:
        def __setattr__(self, name, value):
            raise RuntimeError('security denied')
    closed = []
    composer = SimpleNamespace(_app=DeniedSecurity(), close=lambda **kw: closed.append(kw))
    monkeypatch.setattr(api.word_host, 'create_dedicated_composer', lambda path: composer)
    with pytest.raises(RuntimeError, match='security denied'):
        api._create_owned_composer('writer', tmp_path)
    assert closed == [{'save_changes': False}]


def test_invalid_source_fails_before_any_native_application_creation(api, monkeypatch, tmp_path):
    source = tmp_path / 'source.xlsx'
    source.write_bytes(b'not Office')
    monkeypatch.setattr(api, '_create_owned_composer', lambda *args: pytest.fail('no native launch'))
    with pytest.raises(ValueError):
        api.WindowsExcelSession.open_document(source)
    assert source.read_bytes() == b'not Office'


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_native_save_alias_cannot_escape_private_staging(api, monkeypatch, tmp_path, kind):
    session, composer, source = owned(api, monkeypatch, tmp_path, kind)
    target = tmp_path / ('copy' + source.suffix)
    getattr(session, 'save_' + source.suffix[1:])(target)
    assert target.exists() and Path(composer._doc.FullName).parent == session.staging_root
    session.close()


@pytest.mark.parametrize('kind', ['writer', 'slide'])
def test_structural_image_uses_private_resource_copy(api, monkeypatch, tmp_path, kind):
    session, composer, source = owned(api, monkeypatch, tmp_path, kind)
    image = tmp_path / 'image.png'
    image.write_bytes(b'image bytes')
    seen = []
    composer.apply_structural_op = lambda op: seen.append(op)
    operation = {'op': 'insert', 'type': 'image', 'parent': 'slide:1', 'props': {'path': str(image)}}
    session.apply_structural_op(operation)
    copied = Path(seen[0]['props']['path'])
    assert copied != image and copied.parent == session.staging_root
    assert copied.read_bytes() == image.read_bytes() and operation['props']['path'] == str(image)
    session.close()


def test_workbook_snapshot_does_not_read_foreign_active_window_freeze_state(api, monkeypatch, tmp_path):
    session, composer, source = owned(api, monkeypatch, tmp_path, 'sheet')
    composer._doc.Windows.Item(1).FreezePanes = False
    composer._app.ActiveWindow = SimpleNamespace(FreezePanes=True)
    composer.inspect_document = lambda **kwargs: {'sheets': [{'id': 'sheet:1', 'freeze_panes': True}]}
    assert session.inspect_document()['sheets'][0]['freeze_panes'] is False
    session.close()


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_preflight_current_save_rejects_unsaved_live_document(api, monkeypatch, kind):
    app, doc = attach(api, monkeypatch, kind)
    session = cls(api, kind).attach_active()
    with pytest.raises(ValueError, match='saved'):
        session.preflight_save()
    assert not doc.save_calls and not doc.closed
    session.close()


@pytest.mark.parametrize('kind', ['sheet', 'slide'])
def test_preflight_copy_rejects_non_native_suffix_and_same_live_path(api, monkeypatch, tmp_path, kind):
    app, doc = attach(api, monkeypatch, kind)
    suffix = '.xlsx' if kind == 'sheet' else '.pptx'
    source = package(tmp_path / ('source' + suffix), kind)
    doc.FullName, doc.Path = str(source), str(source.parent)
    session = cls(api, kind).attach_active()
    with pytest.raises(ValueError):
        session.preflight_save(tmp_path / ('output.xls' if kind == 'sheet' else 'output.ppt'))
    with pytest.raises(ValueError, match='differ'):
        session.preflight_save(source, overwrite=True)
    assert not doc.save_calls
    session.close()


def test_preflight_copy_existing_destination_requires_explicit_overwrite(api, monkeypatch, tmp_path):
    app, doc = attach(api, monkeypatch, 'sheet')
    session = api.WindowsExcelSession.attach_active()
    target = package(tmp_path / 'existing.xlsx', 'sheet')
    original = target.read_bytes()
    with pytest.raises(FileExistsError):
        session.preflight_save(target)
    session.preflight_save(target, overwrite=True)
    assert target.read_bytes() == original
    session.close()


@pytest.mark.parametrize('kind', ['writer', 'sheet', 'slide'])
def test_preflight_current_save_accepts_real_native_path_without_saving(api, monkeypatch, tmp_path, kind):
    app, doc = attach(api, monkeypatch, kind)
    source = package(tmp_path / ('source' + {'writer':'.docx', 'sheet':'.xlsx', 'slide':'.pptx'}[kind]), kind)
    doc.FullName, doc.Path = str(source), str(source.parent)
    session = cls(api, kind).attach_active()
    session.preflight_save()
    assert not doc.save_calls and not doc.Saved
    session.close()


def test_preflight_writer_copy_rejects_before_mutation(api, monkeypatch, tmp_path):
    app, doc = attach(api, monkeypatch, 'writer')
    session = api.WindowsWordSession.attach_active()
    with pytest.raises(RuntimeError, match='save-copy'):
        session.preflight_save(tmp_path / 'copy.docx')
    session.close()


class NativeShape:
    def __init__(self, collection, label):
        self.collection, self.label = collection, label
        self.cut_calls = 0
    def Cut(self):
        self.cut_calls += 1
        self.collection.values.remove(self)
    def Copy(self):
        self.collection.clipboard['value'] = self
    def Delete(self):
        self.collection.values.remove(self)
    def Duplicate(self):
        duplicate=NativeShape(self.collection,self.label+' copy')
        self.collection.values.append(duplicate)
        return Collection([duplicate])


class NativeShapes(Collection):
    def __init__(self, clipboard):
        super().__init__()
        self.clipboard=clipboard
        self.fail_paste=False
    def Paste(self):
        if self.fail_paste:
            raise RuntimeError('paste unavailable')
        original=self.clipboard['value']
        duplicate=NativeShape(self,original.label+' copy')
        self.values.append(duplicate)
        return Collection([duplicate])


def live_shape_case(api,monkeypatch):
    app,doc=attach(api,monkeypatch,'slide')
    clipboard={'value':'user clipboard'}
    first,second=NativeShapes(clipboard),NativeShapes(clipboard)
    original=NativeShape(first,'original')
    first.values.append(original)
    doc.Slides.values.extend([SimpleNamespace(Shapes=first),SimpleNamespace(Shapes=second)])
    session=api.WindowsPowerPointSession.attach_active()
    return session,first,second,original,clipboard


def test_native_shape_invalid_destination_is_rejected_before_cut_or_copy(api,monkeypatch):
    s,first,second,original,clipboard=live_shape_case(api,monkeypatch)
    with pytest.raises((ValueError,IndexError)):
        s.apply_structural_op({'op':'move','target':'slide:1/shape:1','to':{'slide':99}})
    assert first.values == [original] and original.cut_calls==0
    assert clipboard['value']=='user clipboard'
    s.close()


def test_native_shape_failed_paste_preserves_original_and_reports_clipboard_change(api,monkeypatch):
    s,first,second,original,clipboard=live_shape_case(api,monkeypatch)
    second.fail_paste=True
    with pytest.raises(RuntimeError,match='paste') as caught:
        s.apply_structural_op({'op':'move','target':'slide:1/shape:1','to':{'slide':2}})
    assert first.values == [original] and second.Count==0
    assert clipboard['value'] is original and caught.value.clipboard_changed is True
    s.close()


def test_native_shape_success_copies_then_deletes_and_reports_clipboard_change(api,monkeypatch):
    s,first,second,original,clipboard=live_shape_case(api,monkeypatch)
    result=s.apply_structural_op({'op':'move','target':'slide:1/shape:1','to':{'slide':2}})
    assert result['moved'] and first.Count==0 and second.Count==1
    assert original.cut_calls==0 and clipboard['value'] is original
    assert result['clipboard_changed'] is True
    s.close()


def test_native_excel_cancelled_delete_does_not_report_success(api,monkeypatch):
    app,doc=attach(api,monkeypatch,'sheet')
    app.DisplayAlerts=False
    doc.Worksheets.values.extend([SimpleNamespace(Delete=lambda:False),SimpleNamespace()])
    s=api.WindowsExcelSession.attach_active()
    with pytest.raises(RuntimeError,match='delet'):
        s.apply_structural_op({'op':'remove','target':'sheet:1'})
    assert doc.Worksheets.Count==2 and app.DisplayAlerts is False
    s.close()


def test_native_excel_shared_delete_refuses_interactive_alert_before_call(api,monkeypatch):
    app,doc=attach(api,monkeypatch,'sheet')
    app.DisplayAlerts=True
    doc.Worksheets.values.extend([SimpleNamespace(Delete=lambda:pytest.fail('would show dialog')),SimpleNamespace()])
    s=api.WindowsExcelSession.attach_active()
    with pytest.raises(RuntimeError,match='interactive'):
        s.apply_structural_op({'op':'remove','target':'sheet:1'})
    assert app.DisplayAlerts is True
    s.close()


class NativeSlide:
    def __init__(self,collection,label):
        self.collection,self.label=collection,label
    @property
    def SlideIndex(self):
        return self.collection.values.index(self)+1
    def MoveTo(self,index):
        if not 1<=index<=self.collection.Count:
            raise ValueError('PowerPoint final index outside collection')
        self.collection.values.remove(self)
        self.collection.values.insert(index-1,self)
    def Duplicate(self):
        duplicate=NativeSlide(self.collection,self.label+' copy')
        self.collection.values.insert(self.SlideIndex,duplicate)
        return Collection([duplicate])
    def Delete(self):
        self.collection.values.remove(self)


@pytest.mark.parametrize('verb',['move','clone'])
def test_native_slide_end_uses_final_index_without_wps_offset(api,monkeypatch,verb):
    app,doc=attach(api,monkeypatch,'slide')
    doc.Slides.values.extend([NativeSlide(doc.Slides,label) for label in ['A','B','C']])
    s=api.WindowsPowerPointSession.attach_active()
    s.apply_structural_op({'op':verb,'target':'slide:1','to':'end'})
    assert [v.label for v in doc.Slides.values] == (['B','C','A'] if verb=='move' else ['A','B','C','A copy'])
    s.close()


def test_native_shape_delete_failure_rolls_back_exact_pasted_copy(api,monkeypatch):
    s,first,second,original,clipboard=live_shape_case(api,monkeypatch)
    original.Delete=lambda:(_ for _ in ()).throw(RuntimeError('source deletion failed'))
    with pytest.raises(RuntimeError,match='source deletion'):
        s.apply_structural_op({'op':'move','target':'slide:1/shape:1','to':{'slide':2}})
    assert first.values==[original] and second.Count==0 and clipboard['value'] is original
    s.close()


def test_native_shape_noop_and_same_slide_duplicate_do_not_use_clipboard(api,monkeypatch):
    s,first,second,original,clipboard=live_shape_case(api,monkeypatch)
    s.apply_structural_op({'op':'move','target':'slide:1/shape:1','to':{'slide':1}})
    assert first.Count==1
    s.apply_structural_op({'op':'clone','target':'slide:1/shape:1','to':{'slide':1}})
    assert first.Count==2 and clipboard['value']=='user clipboard'
    s.close()


@pytest.mark.parametrize('verb',['move','clone'])
@pytest.mark.parametrize('to',[{'before':'slide:3'},{'after':'slide:3'},{'index':2}])
def test_native_slide_anchors_preserve_final_order(api,monkeypatch,verb,to):
    app,doc=attach(api,monkeypatch,'slide')
    doc.Slides.values.extend([NativeSlide(doc.Slides,label) for label in ['A','B','C']])
    s=api.WindowsPowerPointSession.attach_active()
    s.apply_structural_op({'op':verb,'target':'slide:1','to':to})
    expected = {
        ('move','before'):['B','A','C'], ('move','after'):['B','C','A'],('move','index'):['B','A','C'],
        ('clone','before'):['A','B','A copy','C'], ('clone','after'):['A','B','C','A copy'],('clone','index'):['A','A copy','B','C'],
    }[(verb,next(iter(to)))]
    assert [v.label for v in doc.Slides.values]==expected
    s.close()


def test_native_slide_invalid_clone_anchor_does_not_duplicate(api,monkeypatch):
    app,doc=attach(api,monkeypatch,'slide')
    doc.Slides.values.extend([NativeSlide(doc.Slides,label) for label in ['A','B']])
    s=api.WindowsPowerPointSession.attach_active()
    with pytest.raises(ValueError):
        s.apply_structural_op({'op':'clone','target':'slide:1','to':{'after':'slide:99'}})
    assert [v.label for v in doc.Slides.values]==['A','B']
    s.close()


def test_native_private_excel_delete_restores_alert_preference_on_failure(api,monkeypatch,tmp_path):
    s,composer,source=owned(api,monkeypatch,tmp_path,'sheet')
    composer._app.DisplayAlerts=True
    def cancelled():
        assert composer._app.DisplayAlerts is False
        return False
    composer._doc.Worksheets.values.extend([SimpleNamespace(Delete=cancelled),SimpleNamespace()])
    with pytest.raises(RuntimeError,match='delet'):
        s.apply_structural_op({'op':'remove','target':'sheet:1'})
    assert composer._app.DisplayAlerts is True and composer._doc.Worksheets.Count==2
    s.close()


def test_native_clipboard_side_effect_does_not_overwrite_intervening_user_copy(api,monkeypatch):
    s,first,second,original,clipboard=live_shape_case(api,monkeypatch)
    native_paste=second.Paste
    def paste_then_user_copy():
        result=native_paste()
        clipboard['value']='user copied new content during native operation'
        return result
    second.Paste=paste_then_user_copy
    result=s.apply_structural_op({'op':'clone','target':'slide:1/shape:1','to':{'slide':2}})
    assert result['clipboard_changed'] is True
    assert clipboard['value']=='user copied new content during native operation'
    assert first.values==[original] and second.Count==1
    s.close()


class NativeWorksheet:
    def __init__(self,collection,name):
        self.collection,self.Name=collection,name
        self.token=object()
        self.calls=[]
    @property
    def Index(self):
        return self.collection.values.index(self)+1
    def Move(self,*args,**kwargs):
        assert not args and len(kwargs)==1
        key,anchor=next(iter(kwargs.items()))
        self.calls.append(('Move',key,anchor.Name))
        self.collection.values.remove(self)
        index=self.collection.values.index(anchor)+(key=='After')
        self.collection.values.insert(index,self)
    def Copy(self,*args,**kwargs):
        assert not args and len(kwargs)==1
        key,anchor=next(iter(kwargs.items()))
        self.calls.append(('Copy',key,anchor.Name))
        index=self.collection.values.index(anchor)+(key=='After')
        self.collection.values.insert(index,NativeWorksheet(self.collection,self.Name+' copy'))


@pytest.mark.parametrize('to,expected',[({'after':1},['A','A copy','B','C']),({'before':2},['A','A copy','B','C']),('end',['A','B','C','A copy'])])
def test_native_excel_clone_reports_actual_inserted_index_and_uses_keyword_arguments(api,monkeypatch,to,expected):
    app,doc=attach(api,monkeypatch,'sheet')
    doc.Worksheets.values.extend([NativeWorksheet(doc.Worksheets,name) for name in ['A','B','C']])
    s=api.WindowsExcelSession.attach_active()
    result=s.apply_structural_op({'op':'clone','target':'sheet:1','to':to})
    assert [w.Name for w in doc.Worksheets.values]==expected
    assert result['path']=='sheet:'+str(expected.index('A copy')+1)
    s.close()


def test_native_excel_move_end_uses_explicit_after_and_preserves_collection(api,monkeypatch):
    app,doc=attach(api,monkeypatch,'sheet')
    doc.Worksheets.values.extend([NativeWorksheet(doc.Worksheets,name) for name in ['A','B','C']])
    original=doc.Worksheets.Item(1)
    s=api.WindowsExcelSession.attach_active()
    result=s.apply_structural_op({'op':'move','target':'sheet:1','to':'end'})
    assert [w.Name for w in doc.Worksheets.values]==['B','C','A']
    assert original.calls==[('Move','After','C')] and result['path']=='sheet:3'
    s.close()


def test_native_excel_invalid_clone_destination_does_not_create_workbook_or_sheet(api,monkeypatch):
    app,doc=attach(api,monkeypatch,'sheet')
    doc.Worksheets.values.extend([NativeWorksheet(doc.Worksheets,name) for name in ['A','B']])
    s=api.WindowsExcelSession.attach_active()
    with pytest.raises(ValueError):
        s.apply_structural_op({'op':'clone','target':'sheet:1','to':{'after':99}})
    assert doc.Worksheets.Count==2 and not doc.Worksheets.Item(1).calls
    s.close()
