"""Dedicated native Microsoft Word host; never dispatches a WPS ProgID.

The composer inherits the complete WriterComposer native implementation. Only
host acquisition, document opening, saving and cleanup differ. pywin32 is lazy
so routing and tests stay importable on platforms without Microsoft Office.
"""
from __future__ import annotations

from dataclasses import dataclass
import ntpath
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Optional
from uuid import uuid4

from ..writer import WriterComposer


class WordIdentityError(RuntimeError):
    """The COM instance or document cannot safely be attributed to this job."""


def _word_processes(only_pid: Optional[int] = None) -> dict[int, str]:
    import ctypes
    from ctypes import wintypes
    import win32process

    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    result = {}
    for pid in ([only_pid] if only_pid is not None else win32process.EnumProcesses()):
        handle = kernel.OpenProcess(0x1000, False, pid)
        if not handle:
            continue
        try:
            buffer = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(len(buffer))
            if kernel.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                if ntpath.basename(buffer.value).casefold() == 'winword.exe':
                    result[int(pid)] = buffer.value
        finally:
            kernel.CloseHandle(handle)
    return result


def _load_dependencies():
    import pythoncom
    import win32com.client
    import win32process
    import win32gui

    def caption_windows(pid, caption):
        matches = []

        def observe(hwnd, _):
            if (win32process.GetWindowThreadProcessId(hwnd)[1] == pid
                    and win32gui.GetClassName(hwnd) == 'OpusApp'
                    and win32gui.GetWindowText(hwnd) == caption):
                matches.append(hwnd)

        win32gui.EnumWindows(observe, None)
        return matches

    def active():
        try:
            return win32com.client.GetActiveObject('Word.Application')
        except Exception as exc:
            if getattr(exc, 'hresult', None) != -2147221021:  # MK_E_UNAVAILABLE
                raise
            return None

    return SimpleNamespace(
        initialize=pythoncom.CoInitialize,
        uninitialize=pythoncom.CoUninitialize,
        active=active,
        dispatch=lambda: win32com.client.DispatchEx('Word.Application'),
        processes=_word_processes,
        process_image=lambda pid: _word_processes(pid).get(pid),
        identity=lambda obj: obj._oleobj_.QueryInterface(pythoncom.IID_IUnknown),
        window_pid=lambda hwnd: win32process.GetWindowThreadProcessId(hwnd)[1],
        caption_windows=caption_windows,
    )


def _windows_path(value: Any) -> str:
    return ntpath.normcase(ntpath.normpath(str(value)))


@dataclass(frozen=True)
class WordIdentity:
    pid: int
    executable: str


def _empty_application_hwnd(app, pid, deps):
    """Bind empty Word before Add, including versions without Application.Hwnd.

    The caller has already checked new-process isolation, executable, IUnknown,
    and zero documents. A temporary per-instance Caption challenge binds that
    exact COM object to an OpusApp window in the candidate PID. It changes no
    document, security setting, template or running-object registration.
    """
    try:
        return int(app.Hwnd), 'Application.Hwnd'
    except AttributeError:
        pass
    original = str(app.Caption)
    marker = 'WPSComposer-identity-' + uuid4().hex
    try:
        app.Caption = marker
        matches = deps.caption_windows(pid, marker)
        if len(matches) != 1:
            raise WordIdentityError('Empty Word caption did not identify exactly one new-process window')
        hwnd = int(matches[0])
    finally:
        # Never overwrite a concurrent caption change, and never continue to Add
        # if restoring this task's challenge cannot be verified.
        if str(app.Caption) != marker:
            raise WordIdentityError('Empty Word caption changed during identity verification')
        app.Caption = original
        if str(app.Caption) != original:
            raise WordIdentityError('Empty Word caption restoration could not be verified')
    return hwnd, 'restored per-instance Caption challenge'


class NativeWordComposer(WriterComposer):
    """Writer methods operating exclusively on a verified owned Word document."""

    def reset(self):
        # The shared M5 planner fixes A4. Do not inherit a user's Letter or
        # landscape Normal template; conversion does not call this operation.
        self._verify_document(self._doc)
        self.set_orientation(False)
        self.set_page_size(595.28, 841.89)

    @property
    def selection(self):
        # The shared writer API uses Selection primitives. Anchor them to the
        # owned window, since another document may become the app's active one.
        self._verify_document(self._doc)
        selection = self._doc.Windows.Item(1).Selection
        if self._deps.identity(selection.Document) != self._deps.identity(self._doc):
            raise WordIdentityError('Word selection is not bound to the owned document')
        return selection

    def _verify_application(self) -> None:
        observed = self._deps.process_image(self.identity.pid)
        if (
            not observed
            or _windows_path(observed) != _windows_path(self.identity.executable)
            or self._deps.identity(self._app) != self._application_token
            or str(self._app.Name) != 'Microsoft Word'
            or _windows_path(self._app.Path) != _windows_path(ntpath.dirname(observed))
        ):
            raise WordIdentityError('Native Word process identity changed; refusing mutation or Quit')

    def _verify_document(self, doc) -> None:
        self._verify_application()
        if self._deps.identity(doc.Application) != self._application_token:
            raise WordIdentityError('Document application is not the dedicated Word instance')
        if int(doc.Windows.Count) < 1:
            raise WordIdentityError('Owned Word document has no verifiable window')
        hwnd = int(doc.Windows.Item(1).Hwnd)
        if not hwnd or self._deps.window_pid(hwnd) != self.identity.pid:
            raise WordIdentityError('Document HWND does not belong to the new WINWORD.EXE process')

    def _require_staged_path(self, value) -> Path:
        path = Path(value).expanduser().resolve()
        if not self._staging_dir:
            raise WordIdentityError('An owned staging directory is required to open documents')
        root = Path(self._staging_dir).resolve()
        if root not in path.parents or path.is_symlink():
            raise WordIdentityError('Document path is outside the owned staging directory')
        return path

    def open_existing_for_patch(self, path):
        return self.open_owned_document(path, read_only=False)

    def open_owned_document(self, path, *, read_only=True):
        source = self._require_staged_path(path)
        self._verify_document(self._doc)
        self._doc.Close(SaveChanges=0)
        self._doc = None
        # No macro/security preferences are changed. Source files belong to this
        # job's private directory, never the user's original document.
        doc = self._app.Documents.Open(
            FileName=str(source), ConfirmConversions=False, ReadOnly=read_only,
            AddToRecentFiles=False, Visible=False,
        )
        self._verify_document(doc)
        if _windows_path(doc.FullName) != _windows_path(source):
            raise WordIdentityError('Reopened Word document path does not match the owned source')
        self._doc = doc
        self._read_only = bool(read_only)
        return self

    def save_docx(self, path):
        self._verify_document(self._doc)
        target = self._require_staged_path(path)
        self._doc.SaveAs2(FileName=str(target), FileFormat=16, AddToRecentFiles=False)
        if _windows_path(self._doc.FullName) != _windows_path(target):
            raise WordIdentityError('Word saved document path does not match the owned target')
        return str(target)

    def export_pdf(self, path):
        self._verify_document(self._doc)
        target = self._require_staged_path(path)
        self._doc.ExportAsFixedFormat(
            OutputFileName=str(target), ExportFormat=17, OpenAfterExport=False,
        )
        return str(target)

    def close(self, save_changes=False):
        if self._closed:
            return
        try:
            if self._doc is not None:
                self._verify_document(self._doc)
                self._doc.Close(SaveChanges=-1 if save_changes else 0)
                self._doc = None
            self._verify_application()
            if int(self._app.Documents.Count) != 0:
                raise WordIdentityError('Unexpected documents appeared; dedicated Word left open')
            self._app.Quit(SaveChanges=0)
            self._closed = True
            self._app = None
        except BaseException as exc:
            self.cleanup_error = exc
            raise
        finally:
            if self._com_initialized:
                self._deps.uninitialize()
                self._com_initialized = False


def create_dedicated_composer(staging_dir: Optional[str] = None) -> NativeWordComposer:
    """Verify new WINWORD process, IUnknown isolation, empty host and owned HWND.

    Failure before the document HWND gate never grants mutation/cleanup rights.
    Such uncertain instances are intentionally retained for manual inspection.
    """
    deps = _load_dependencies()
    deps.initialize()
    try:
        previous = deps.active()
        before = deps.processes()
        app = deps.dispatch()
        candidates = {pid: path for pid, path in deps.processes().items() if pid not in before}
        if len(candidates) != 1:
            raise WordIdentityError('Cannot identify exactly one new WINWORD.EXE process')
        pid, executable = next(iter(candidates.items()))
        if (
            ntpath.basename(executable).casefold() != 'winword.exe'
            or str(app.Name) != 'Microsoft Word'
            or _windows_path(app.Path) != _windows_path(ntpath.dirname(executable))
        ):
            raise WordIdentityError('COM application is not the new native Microsoft Word process')
        token = deps.identity(app)
        if previous is not None and token == deps.identity(previous):
            raise WordIdentityError('DispatchEx returned the preexisting Word application')
        if int(app.Documents.Count) != 0:
            raise WordIdentityError('DispatchEx application already contains documents')
        # Bind the empty application itself before Documents.Add can mutate it.
        # The document HWND remains a second gate once the owned doc exists.
        application_hwnd, binding_method = _empty_application_hwnd(app, pid, deps)
        if not application_hwnd or deps.window_pid(application_hwnd) != pid:
            raise WordIdentityError('Application HWND does not belong to the new WINWORD.EXE process')
        if deps.identity(app) != token or int(app.Documents.Count) != 0:
            raise WordIdentityError('Empty Word application changed before Documents.Add')
        composer = NativeWordComposer.__new__(NativeWordComposer)
        composer._deps = deps
        composer.identity = WordIdentity(pid=int(pid), executable=executable)
        composer.application_binding = {'hwnd': application_hwnd, 'pid': int(pid),
                                        'method': binding_method, 'documents_before_add': 0}
        composer._application_token = token
        composer._app = app
        composer._doc = None
        composer._staging_dir = staging_dir
        composer._path = None
        composer._read_only = False
        composer._visible = False
        composer._owns_app = True
        composer._owns_doc = True
        composer._com_initialized = True
        composer._first_section_configured = False
        composer._closed = False
        composer.cleanup_error = None
        doc = app.Documents.Add()
        composer._verify_document(doc)
        composer._doc = doc
        return composer
    except BaseException:
        deps.uninitialize()
        raise


# WindowsLongformExecutor uses these opt-outs to retain native failure evidence
# and prevent a retry from creating another host after content mutation begins.
create_dedicated_composer.allow_host_retry = False
create_dedicated_composer.strict_cleanup = True
