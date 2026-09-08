"""Native COM boundary doubles; never start Excel or PowerPoint."""
from pathlib import Path
from types import SimpleNamespace

import pytest


class Collection:
    def __init__(self, values=()):
        self.values = list(values)

    @property
    def Count(self):
        return len(self.values)

    def Item(self, index):
        return self.values[index - 1]

    __call__ = Item


class Document:
    def __init__(self, app, name='owned', hwnd=44):
        self.Application = app
        self.Name = self.FullName = name
        self.Saved = False
        self.token = object()
        self.Windows = Collection([SimpleNamespace(Hwnd=hwnd)])
        self.closed = False
        self.exports = []
        self.calculations = 0
        def calculate():
            self.calculations += 1
        self.Worksheets = Collection([SimpleNamespace(Name='Sheet1', Calculate=calculate)])
        self.Slides = Collection()

    def Close(self, **kwargs):
        if self.Application.component == 'spreadsheet':
            assert kwargs == {'SaveChanges': False}
        else:
            assert kwargs == {} and self.Saved
        self.closed = True
        self.Application.documents.values.remove(self)

    def SaveAs(self, **kwargs):
        self.FullName = kwargs['Filename'] if 'Filename' in kwargs else kwargs['FileName']
        self.saved_args = kwargs

    def ExportAsFixedFormat(self, **kwargs):
        self.exports.append(kwargs)


class Documents(Collection):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.added = 0
        self.opened = []

    def Add(self, **kwargs):
        self.added += 1
        self.add_arguments = kwargs
        doc = Document(self.app)
        self.values.append(doc)
        return doc

    def Open(self, **kwargs):
        self.opened.append(kwargs)
        doc = Document(self.app, kwargs['Filename'] if 'Filename' in kwargs else kwargs['FileName'])
        self.values.append(doc)
        return doc


class App:
    Path = r'C:\Program Files\Microsoft Office\root\Office16'
    Hwnd = 43
    Version = '16.0'

    def __init__(self, component):
        self.component = component
        self.Name = 'Microsoft Excel' if component == 'spreadsheet' else 'Microsoft PowerPoint'
        self.token = object()
        self.documents = Documents(self)
        self.Workbooks = self.Presentations = self.documents
        self.AutomationSecurity = 1
        self.quit = False

    def Quit(self):
        self.quit = True


@pytest.fixture
def host():
    from skills.WPSComposer.scripts.msoffice import windows_office_host
    return windows_office_host


def setup_host(host, monkeypatch, component, *, shared=False):
    app = App(component)
    executable = App.Path + ('\\EXCEL.EXE' if component == 'spreadsheet' else '\\POWERPNT.EXE')
    inventory = iter([{22: executable} if shared else {}, {22: executable}])
    def scan():
        return next(inventory, {22: executable})
    deps = SimpleNamespace(
        initialize=lambda: None, uninitialize=lambda: None,
        dispatch=lambda progid: app, processes=scan,
        process_image=lambda pid: executable if pid == 22 else None,
        identity=lambda obj: obj.token, window_pid=lambda hwnd: 22,
        powerpoint_windows=lambda: [(43, SimpleNamespace(Application=app, Presentation=doc))
                                    for doc in (app.documents.values or [None])],
    )
    monkeypatch.setattr(host, '_load_dependencies', lambda component: deps)
    return app, deps


@pytest.mark.parametrize('component', ['spreadsheet', 'presentation'])
@pytest.mark.parametrize('damage', ['wrong_name', 'wrong_image', 'wrong_path', 'wrong_binding'])
def test_wrong_application_never_creates_or_closes_documents(host, monkeypatch, tmp_path, component, damage):
    app, deps = setup_host(host, monkeypatch, component)
    if damage == 'wrong_name':
        app.Name = 'WPS Office'
    elif damage == 'wrong_image':
        deps.process_image = lambda pid: App.Path + '\\wps.exe'
    elif damage == 'wrong_path':
        app.Path = r'C:\Other'
    elif component == 'spreadsheet':
        deps.window_pid = lambda hwnd: 999
    else:
        deps.powerpoint_windows = lambda: [(43, SimpleNamespace(Application=App(component)))]
    with pytest.raises(host.OfficeIdentityError):
        host.create_composer(component, tmp_path)
    assert app.documents.added == 0 and not app.quit


@pytest.mark.parametrize('component', ['spreadsheet', 'presentation'])
def test_shared_host_preserves_unsaved_sentinel_and_security(host, monkeypatch, tmp_path, component):
    app, _ = setup_host(host, monkeypatch, component, shared=True)
    sentinel = Document(app, 'unsaved user sentinel')
    app.documents.values.append(sentinel)
    composer = host.create_composer(component, tmp_path)
    owned = composer._doc
    composer.close()
    assert owned.closed and app.documents.values == [sentinel]
    assert not sentinel.closed and not sentinel.Saved
    assert not app.quit and app.AutomationSecurity == 1


def test_excel_private_host_uses_single_sheet_template_and_exact_save_export(host, monkeypatch, tmp_path):
    app, _ = setup_host(host, monkeypatch, 'spreadsheet')
    composer = host.create_composer('spreadsheet', tmp_path)
    doc = composer._doc
    assert app.documents.add_arguments == {'Template': -4167}
    assert app.AutomationSecurity == 3
    composer.save_xlsx(tmp_path / 'out.xlsx')
    composer.export_pdf(tmp_path / 'out.pdf')
    assert doc.saved_args['FileFormat'] == 51
    assert doc.exports == [{'Type': 0, 'Filename': str(tmp_path / 'out.pdf'), 'OpenAfterPublish': False}]
    assert doc.calculations == 2
    composer.close()
    assert app.quit and composer._closed


def test_powerpoint_binds_native_window_not_word_hwnd_or_caption(host, monkeypatch, tmp_path):
    app, deps = setup_host(host, monkeypatch, 'presentation', shared=True)
    delattr(App, 'Hwnd')
    try:
        composer = host.create_composer('presentation', tmp_path)
        doc = composer._doc
        composer.save_pptx(tmp_path / 'out.pptx')
        composer.export_pdf(tmp_path / 'out.pdf')
        assert doc.saved_args['FileFormat'] == 24
        assert doc.exports == [{'Path': str(tmp_path / 'out.pdf'), 'FixedFormatType': 2}]
        composer.close()
        assert not app.quit
    finally:
        App.Hwnd = 43


def test_unbound_empty_powerpoint_fails_before_add(host, monkeypatch, tmp_path):
    app, deps = setup_host(host, monkeypatch, 'presentation')
    deps.powerpoint_windows = lambda: []
    with pytest.raises(host.OfficeIdentityError):
        host.create_composer('presentation', tmp_path)
    assert app.documents.added == 0 and not app.quit


@pytest.mark.parametrize('component', ['spreadsheet', 'presentation'])
def test_new_document_identity_cannot_alias_sentinel(host, monkeypatch, tmp_path, component):
    app, deps = setup_host(host, monkeypatch, component, shared=True)
    sentinel = Document(app, 'sentinel')
    app.documents.values.append(sentinel)
    app.documents.Add = lambda **kwargs: sentinel
    with pytest.raises(host.OfficeIdentityError):
        host.create_composer(component, tmp_path)
    assert not sentinel.closed and not app.quit


def test_foreign_document_blocks_quit_without_closing_it(host, monkeypatch, tmp_path):
    app, _ = setup_host(host, monkeypatch, 'spreadsheet')
    composer = host.create_composer('spreadsheet', tmp_path)
    foreign = Document(app, 'concurrent user file')
    app.documents.values.append(foreign)
    with pytest.raises(host.OfficeIdentityError):
        composer.close()
    assert not foreign.closed and not app.quit and composer.cleanup_error


@pytest.mark.parametrize('component', ['spreadsheet', 'presentation'])
def test_source_path_outside_staging_rejected_before_any_close(host, monkeypatch, tmp_path, component):
    app, _ = setup_host(host, monkeypatch, component)
    root = tmp_path / 'owned'
    root.mkdir()
    composer = host.create_composer(component, root)
    initial = composer._doc
    with pytest.raises(host.OfficeIdentityError):
        composer.open_owned_document(tmp_path / 'user.xlsx')
    assert not initial.closed and not app.documents.opened
    composer.close()


def test_no_wps_fallback_after_com_activation_failure(host, monkeypatch, tmp_path):
    app, deps = setup_host(host, monkeypatch, 'spreadsheet')
    attempts = []
    def dispatch(progid):
        attempts.append(progid)
        raise RuntimeError('activation failed')
    deps.dispatch = dispatch
    with pytest.raises(RuntimeError, match='activation failed'):
        host.create_composer('spreadsheet', tmp_path)
    assert attempts == ['Excel.Application']


@pytest.mark.parametrize('class_name', ['NativeSheetComposer', 'NativeSlideComposer'])
def test_legacy_entrypoints_cannot_bypass_native_identity_factory(host, class_name):
    cls = getattr(host, class_name)
    with pytest.raises(host.OfficeIdentityError):
        cls()
    with pytest.raises(host.OfficeIdentityError):
        cls.attach_active()
    with pytest.raises(host.OfficeIdentityError):
        cls.open_document('user.xlsx')


@pytest.mark.parametrize('component,fmt', [('spreadsheet', 'xlsx'), ('presentation', 'pptx')])
def test_native_open_is_read_only_and_rejects_macro_package_before_close(host, monkeypatch, tmp_path, component, fmt):
    import zipfile
    app, _ = setup_host(host, monkeypatch, component, shared=True)
    composer = host.create_composer(component, tmp_path)
    initial = composer._doc
    malicious = tmp_path / ('input.' + fmt)
    with zipfile.ZipFile(malicious, 'w') as archive:
        archive.writestr('word/vbaProject.bin', b'nonexecutable test')
    with pytest.raises(ValueError):
        composer.open_owned_document(malicious)
    assert not initial.closed and not app.documents.opened
    composer.close()
