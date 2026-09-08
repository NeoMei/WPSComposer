from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.msoffice import windows_host as host


class Collection:
    def __init__(self, items=()):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def Item(self, index):
        return self.items[index - 1]


class Document:
    def __init__(self, app, hwnd=44, path='Document1'):
        self.Application = app
        self.Windows = Collection([SimpleNamespace(Hwnd=hwnd)])
        self.FullName = path
        self.closed = False

    def Close(self, SaveChanges=0):
        self.closed = True
        self.Application.Documents.items.remove(self)

    def SaveAs2(self, **kwargs):
        self.FullName = kwargs['FileName']


class Documents(Collection):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.added = 0

    def Add(self):
        self.added += 1
        doc = Document(self.app)
        self.items.append(doc)
        return doc

    def Open(self, **kwargs):
        doc = Document(self.app, path=kwargs['FileName'])
        self.items.append(doc)
        return doc


class App:
    Name = 'Microsoft Word'
    Hwnd = 43
    Path = r'C:\Program Files\Microsoft Office\root\Office16'

    def __init__(self, token=1):
        self.token = token
        self.Documents = Documents(self)
        self.quit = False

    def Quit(self, SaveChanges=0):
        self.quit = True


def setup_host(monkeypatch, *, previous=None, new=None, pid=22, inventory=None):
    app = new or App()
    processes = iter(inventory or [{11: App.Path + r'\WINWORD.EXE'}, {11: App.Path + r'\WINWORD.EXE', 22: App.Path + r'\WINWORD.EXE'}])
    latest = {}

    def scan():
        nonlocal latest
        latest = next(processes, latest)
        return latest

    deps = SimpleNamespace(
        active=lambda: previous, dispatch=lambda: app, processes=scan,
        process_image=lambda pid: scan().get(pid),
        window_pid=lambda hwnd: pid, identity=lambda obj: obj.token,
        initialize=lambda: None, uninitialize=lambda: None,
    )
    monkeypatch.setattr(host, '_load_dependencies', lambda: deps)
    return app


@pytest.mark.parametrize('mode', ['shared', 'populated', 'wrong_process', 'wrong_path', 'no_new_process'])
def test_rejects_unverified_application_without_creating_document_or_quitting(monkeypatch, mode):
    app = App()
    previous = App(token=1) if mode == 'shared' else None
    inventory = None
    if mode == 'populated':
        app.Documents.items.append(Document(app))
    if mode == 'wrong_process':
        app.Name = 'WPS Office'
    if mode == 'wrong_path':
        app.Path = r'C:\OtherOffice'
        inventory = [{}, {22: r'C:\Program Files\Microsoft Office\WINWORD.EXE'}]
    if mode == 'no_new_process':
        inventory = [{22: App.Path + r'\WINWORD.EXE'}] * 2
    setup_host(monkeypatch, previous=previous, new=app, inventory=inventory)
    with pytest.raises(host.WordIdentityError):
        host.create_dedicated_composer()
    assert app.Documents.added == 0
    assert not app.quit


def test_application_pid_mismatch_never_mutates_or_quits(monkeypatch):
    app = setup_host(monkeypatch, pid=999)
    with pytest.raises(host.WordIdentityError):
        host.create_dedicated_composer()
    assert app.Documents.added == 0
    assert not app.quit


def test_document_pid_mismatch_retains_uncertain_document(monkeypatch):
    app = setup_host(monkeypatch)
    deps = host._load_dependencies()
    deps.window_pid = lambda hwnd: 22 if hwnd == app.Hwnd else 999
    with pytest.raises(host.WordIdentityError):
        host.create_dedicated_composer()
    assert len(app.Documents.items) == 1
    assert not app.Documents.items[0].closed
    assert not app.quit


def test_verified_composer_preserves_writer_methods_and_closes_only_owned_host(monkeypatch, tmp_path):
    from skills.WPSComposer.scripts.writer import WriterComposer
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer(str(tmp_path))
    assert isinstance(composer, WriterComposer)
    assert callable(composer.pagination_map_for_ranges)
    assert callable(composer.add_semantic_table_native)
    composer.save_docx(str(tmp_path / 'result.docx'))
    composer.close()
    assert app.Documents.Count == 0
    assert app.quit
    assert composer.cleanup_error is None


def test_generation_reset_uses_planned_a4_instead_of_normal_template(monkeypatch, tmp_path):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer(str(tmp_path))
    composer._doc.PageSetup = SimpleNamespace(PageWidth=612, PageHeight=792, Orientation=1)
    composer.reset()
    setup = composer._doc.PageSetup
    assert (setup.PageWidth, setup.PageHeight, setup.Orientation) == (595.28, 841.89, 0)
    composer.close()


def test_unexpected_document_prevents_application_quit(monkeypatch):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer()
    foreign = Document(app, hwnd=55)
    app.Documents.items.append(foreign)
    with pytest.raises(host.WordIdentityError):
        composer.close()
    assert not app.quit
    assert not foreign.closed
    assert composer.cleanup_error is not None


def test_reopened_document_rechecks_hwnd_before_patch(monkeypatch, tmp_path):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer(str(tmp_path))
    source = tmp_path / 'source.docx'
    source.write_bytes(b'owned')
    composer._deps.window_pid = lambda hwnd: 999
    with pytest.raises(host.WordIdentityError):
        composer.open_existing_for_patch(str(source))
    assert not app.quit


def test_cleanup_failure_is_observable_and_does_not_quit(monkeypatch):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer()

    def fail_close(**kwargs):
        raise OSError('Document is busy')

    composer._doc.Close = fail_close
    with pytest.raises(OSError, match='busy'):
        composer.close()
    assert composer.cleanup_error is not None
    assert not app.quit


def test_reopen_rejects_source_outside_owned_root_before_closing(monkeypatch, tmp_path):
    app = setup_host(monkeypatch)
    stage = tmp_path / 'stage'
    stage.mkdir()
    composer = host.create_dedicated_composer(str(stage))
    initial = composer._doc
    with pytest.raises(host.WordIdentityError, match='outside'):
        composer.open_owned_document(tmp_path / 'user-document.docx')
    assert not initial.closed
    assert not app.quit
    composer.close()


def test_reopen_document_path_mismatch_retains_host(monkeypatch, tmp_path):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer(str(tmp_path))
    original_open = app.Documents.Open

    def wrong_path(**kwargs):
        doc = original_open(**kwargs)
        doc.FullName = 'unrelated.docx'
        return doc

    app.Documents.Open = wrong_path
    with pytest.raises(host.WordIdentityError, match='path'):
        composer.open_existing_for_patch(tmp_path / 'owned.docx')
    with pytest.raises(host.WordIdentityError, match='Unexpected'):
        composer.close()
    assert not app.quit
    assert app.Documents.Count == 1


def test_writer_selection_is_bound_to_owned_window_not_active_document(monkeypatch):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer()
    composer._doc.token = 77
    owned_selection = SimpleNamespace(Document=composer._doc)
    composer._doc.Windows.Item(1).Selection = owned_selection
    app.Selection = SimpleNamespace(Document=Document(app, path='User document'))
    assert composer.selection is owned_selection
    composer.close()


def test_selection_rejects_window_bound_to_unrelated_document(monkeypatch):
    app = setup_host(monkeypatch)
    composer = host.create_dedicated_composer()
    composer._doc.token = 77
    foreign = Document(app)
    foreign.token = 88
    selection = SimpleNamespace(Document=foreign)
    composer._doc.Windows.Item(1).Selection = selection
    app.Selection = selection
    with pytest.raises(host.WordIdentityError, match='selection'):
        _ = composer.selection
    composer.close()


class EmptyNativeApp(App):
    Caption = 'Word'

    @property
    def Hwnd(self):
        raise AttributeError('Word.Application.Hwnd')


def test_empty_word_without_application_hwnd_is_bound_before_document_add(monkeypatch):
    app = setup_host(monkeypatch, new=EmptyNativeApp())
    observations = []

    def matching_windows(pid, caption):
        observations.append((pid, app.Documents.Count, app.Caption))
        assert app.Caption == caption and caption != 'Word'
        return [43]

    host._load_dependencies().caption_windows = matching_windows
    composer = host.create_dedicated_composer()
    assert len(observations) == 1 and observations[0][:2] == (22, 0)
    assert app.Caption == 'Word'
    assert app.Documents.added == 1
    composer.close()


@pytest.mark.parametrize('matches', [[], [43, 45]])
def test_caption_proof_failure_restores_caption_without_document_or_quit(monkeypatch, matches):
    app = setup_host(monkeypatch, new=EmptyNativeApp())
    observed = []
    host._load_dependencies().caption_windows = lambda pid, marker: observed.append(marker) or matches
    with pytest.raises(host.WordIdentityError, match='caption'):
        host.create_dedicated_composer()
    assert len(observed) == 1
    assert app.Caption == 'Word'
    assert app.Documents.added == 0 and not app.quit


def test_caption_proof_rechecks_empty_application_before_add(monkeypatch):
    app = setup_host(monkeypatch, new=EmptyNativeApp())

    def matching_windows(pid, marker):
        app.Documents.items.append(Document(app, path='Concurrent user document'))
        return [43]

    host._load_dependencies().caption_windows = matching_windows
    with pytest.raises(host.WordIdentityError):
        host.create_dedicated_composer()
    assert app.Caption == 'Word'
    assert app.Documents.added == 0 and not app.quit


def test_caption_enumeration_failure_restores_original_caption(monkeypatch):
    app = setup_host(monkeypatch, new=EmptyNativeApp())
    app.Caption = 'Original instance caption'

    def unavailable(pid, marker):
        raise OSError('Native window enumeration failed')

    host._load_dependencies().caption_windows = unavailable
    with pytest.raises(OSError, match='enumeration'):
        host.create_dedicated_composer()
    assert app.Caption == 'Original instance caption'
    assert app.Documents.added == 0 and not app.quit


def test_concurrent_caption_change_is_not_overwritten_or_used_for_creation(monkeypatch):
    app = setup_host(monkeypatch, new=EmptyNativeApp())

    def changed(pid, marker):
        app.Caption = 'Concurrent caption'
        return [43]

    host._load_dependencies().caption_windows = changed
    with pytest.raises(host.WordIdentityError, match='caption changed'):
        host.create_dedicated_composer()
    assert app.Caption == 'Concurrent caption'
    assert app.Documents.added == 0 and not app.quit
