"""Verified native Microsoft Excel/PowerPoint hosts, with no WPS fallback.

PowerPoint has no assumed Word-style HWND/Caption API. Its native document
windows are bound through OBJID_NATIVEOM and Application IUnknown. An empty
PowerPoint with no accessible pane fails closed before creating a document.
Shared hosts own only their newly created document and are never quit.
"""
from __future__ import annotations

import ntpath
from pathlib import Path
from types import SimpleNamespace

from ..sheet import SheetComposer
from ..slide import SlideComposer


APPS = {
    'spreadsheet': ('Excel.Application', 'Microsoft Excel', 'EXCEL.EXE', 'Workbooks', 'xlsx'),
    'presentation': ('PowerPoint.Application', 'Microsoft PowerPoint', 'POWERPNT.EXE', 'Presentations', 'pptx'),
}


class OfficeIdentityError(RuntimeError):
    pass


def _load_dependencies(component):
    import ctypes
    from ctypes import wintypes
    import pythoncom
    import win32com.client
    import win32gui
    import win32process

    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                               wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    expected = APPS[component][2].casefold()

    def process_image(pid):
        handle = kernel.OpenProcess(0x1000, False, pid)
        if not handle:
            return None
        try:
            buffer = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(len(buffer))
            if kernel.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
                return buffer.value
            return None
        finally:
            kernel.CloseHandle(handle)

    def processes():
        result = {}
        for pid in win32process.EnumProcesses():
            path = process_image(pid)
            if path and ntpath.basename(path).casefold() == expected:
                result[int(pid)] = path
        return result

    def powerpoint_windows():
        # https://learn.microsoft.com/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow
        # paneClassDC exposes PowerPoint DocumentWindow, not Word Window.
        oleacc = ctypes.OleDLL('oleacc')
        oleacc.AccessibleObjectFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD,
                                                     ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)]
        oleacc.AccessibleObjectFromWindow.restype = ctypes.c_long
        import uuid
        iid = (ctypes.c_ubyte * 16).from_buffer_copy(uuid.UUID(str(pythoncom.IID_IDispatch)).bytes_le)
        handles = []
        def child(hwnd, _):
            if win32gui.GetClassName(hwnd) == 'paneClassDC':
                handles.append(hwnd)
        def top(hwnd, _):
            pid = win32process.GetWindowThreadProcessId(hwnd)[1]
            image = process_image(pid)
            if image and ntpath.basename(image).casefold() == 'powerpnt.exe':
                win32gui.EnumChildWindows(hwnd, child, None)
        win32gui.EnumWindows(top, None)
        result = []
        for hwnd in handles:
            pointer = ctypes.c_void_p()
            try:
                oleacc.AccessibleObjectFromWindow(hwnd, 0xfffffff0, ctypes.byref(iid), ctypes.byref(pointer))
                if pointer.value:
                    # ObjectFromAddress queries/AddRefs its own interface.
                    obj = win32com.client.Dispatch(pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch))
                    result.append((hwnd, obj))
            except (OSError, pythoncom.com_error):
                continue
            finally:
                if pointer.value:
                    table = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
                    ctypes.WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)(table[2])(pointer)
        return result

    return SimpleNamespace(
        initialize=pythoncom.CoInitialize, uninitialize=pythoncom.CoUninitialize,
        dispatch=win32com.client.DispatchEx, processes=processes, process_image=process_image,
        identity=lambda obj: obj._oleobj_.QueryInterface(pythoncom.IID_IUnknown),
        window_pid=lambda hwnd: win32process.GetWindowThreadProcessId(hwnd)[1],
        powerpoint_windows=powerpoint_windows,
    )


def _path(value):
    return ntpath.normcase(ntpath.normpath(str(value)))


def _snapshot(collection, deps):
    return [(deps.identity(doc), str(doc.Name), str(doc.FullName), bool(doc.Saved))
            for doc in (collection.Item(i) for i in range(1, int(collection.Count) + 1))]


def _application_pid(app, component, deps):
    if component == 'spreadsheet':
        hwnd = int(app.Hwnd)
        if not hwnd:
            raise OfficeIdentityError('Excel has no verifiable application HWND')
        return deps.window_pid(hwnd)
    token = deps.identity(app)
    pids = {deps.window_pid(hwnd) for hwnd, window in deps.powerpoint_windows()
            if deps.identity(window.Application) == token}
    if len(pids) != 1:
        raise OfficeIdentityError('PowerPoint native window cannot bind exactly one application process')
    return pids.pop()


class _OwnedOffice:
    def __init__(self, *args, **kwargs):
        raise OfficeIdentityError('Use the explicit create_composer identity factory')

    @classmethod
    def attach_active(cls, *args, **kwargs):
        raise OfficeIdentityError('Active Office attachment requires a dedicated session factory')

    @classmethod
    def open_document(cls, *args, **kwargs):
        raise OfficeIdentityError('Open Office files through the verified private staging factory')

    def __enter__(self):
        self._verify_document()
        return self

    def _verify_application(self):
        image = self._deps.process_image(self.pid)
        if (not image or _path(image) != _path(self.executable)
                or ntpath.basename(image).casefold() != APPS[self.component][2].casefold()
                or str(self._app.Name) != APPS[self.component][1]
                or _path(self._app.Path) != _path(ntpath.dirname(image))
                or self._deps.identity(self._app) != self._application_token):
            raise OfficeIdentityError('Native Office application/process identity changed')

    def _verify_document(self):
        self._verify_application()
        if (self._doc is None or self._deps.identity(self._doc) != self._document_token
                or self._deps.identity(self._doc.Application) != self._application_token):
            raise OfficeIdentityError('Office document is not the exact task-owned object')
        if self.component == 'spreadsheet':
            if self._doc.Windows.Count < 1 or self._deps.window_pid(int(self._doc.Windows.Item(1).Hwnd)) != self.pid:
                raise OfficeIdentityError('Workbook window does not belong to verified EXCEL.EXE')
        else:
            matches = [hwnd for hwnd, window in self._deps.powerpoint_windows()
                       if self._deps.identity(window.Application) == self._application_token
                       and window.Presentation is not None
                       and self._deps.identity(window.Presentation) == self._document_token
                       and self._deps.window_pid(hwnd) == self.pid]
            if not matches:
                raise OfficeIdentityError('Presentation has no matching native DocumentWindow')

    def _staged_path(self, value):
        raw = Path(value).expanduser()
        path = raw.resolve()
        if raw.is_symlink() or self._staging_dir not in path.parents:
            raise OfficeIdentityError('Office file is outside its private staging directory')
        return path

    def _close_document(self):
        self._verify_document()
        if self.component == 'spreadsheet':
            self._doc.Close(SaveChanges=False)
        else:
            # Presentation.Close has no SaveChanges argument. Only the exact
            # owned presentation gets this discard flag; never touch siblings.
            self._doc.Saved = True
            self._doc.Close()
        self._doc = None
        self._document_token = None
        self._ws = None

    def open_owned_document(self, value, *, read_only=True):
        source = self._staged_path(value)
        from .input_validation import validate_native_input
        validate_native_input(source, self.component)
        self._close_document()
        before = _snapshot(self._collection, self._deps)
        if self._owns_app:
            self._app.AutomationSecurity = 3  # ForceDisable, only our isolated app
        if self.component == 'spreadsheet':
            doc = self._collection.Open(Filename=str(source), UpdateLinks=0, ReadOnly=bool(read_only),
                                        AddToMru=False, IgnoreReadOnlyRecommended=True)
        else:
            doc = self._collection.Open(FileName=str(source), ReadOnly=-1 if read_only else 0,
                                        Untitled=0, WithWindow=-1)
        token = self._deps.identity(doc)
        if token in [row[0] for row in before] or _path(doc.FullName) != _path(source):
            raise OfficeIdentityError('Opened Office document is not the private source copy')
        self._doc, self._document_token = doc, token
        self._verify_document()
        self._read_only = bool(read_only)
        return self

    def close(self, save_changes=False):
        if self._closed:
            return
        try:
            if save_changes:
                raise ValueError('Native jobs close only without saving; save explicitly to staging')
            if self._doc is not None:
                self._close_document()
            self._verify_application()
            if _snapshot(self._collection, self._deps) != self._baseline:
                raise OfficeIdentityError('Unrelated Office documents changed; application left open')
            if self._owns_app:
                if self._collection.Count != 0:
                    raise OfficeIdentityError('Unexpected documents prevent application Quit')
                self._app.Quit()
            self._closed = True
            self._app = None
        except BaseException as exc:
            self.cleanup_error = exc
            raise
        finally:
            if self._com_initialized:
                self._deps.uninitialize()
                self._com_initialized = False


class NativeSheetComposer(_OwnedOffice, SheetComposer):
    def _calculate_owned(self):
        # Application.Calculate would also mutate unrelated open workbooks.
        # Refresh only this job's sheets, including under manual calculation.
        for index in range(1, self._doc.Worksheets.Count + 1):
            self._doc.Worksheets.Item(index).Calculate()

    def save_xlsx(self, path):
        self._verify_document()
        target = self._staged_path(path)
        self._calculate_owned()
        self._doc.SaveAs(Filename=str(target), FileFormat=51, AddToMru=False)
        if _path(self._doc.FullName) != _path(target):
            raise OfficeIdentityError('Excel saved outside requested staging path')
        return str(target)

    def export_pdf(self, path):
        self._verify_document()
        target = self._staged_path(path)
        self._calculate_owned()
        self._doc.ExportAsFixedFormat(Type=0, Filename=str(target), OpenAfterPublish=False)
        return str(target)


class NativeSlideComposer(_OwnedOffice, SlideComposer):
    def save_pptx(self, path):
        self._verify_document()
        target = self._staged_path(path)
        self._doc.SaveAs(FileName=str(target), FileFormat=24)
        if _path(self._doc.FullName) != _path(target):
            raise OfficeIdentityError('PowerPoint saved outside requested staging path')
        return str(target)

    def export_pdf(self, path):
        self._verify_document()
        target = self._staged_path(path)
        self._doc.ExportAsFixedFormat(Path=str(target), FixedFormatType=2)
        return str(target)


def create_composer(component, staging_dir):
    if component not in APPS:
        raise ValueError('Unsupported native Office component')
    root = Path(staging_dir).resolve()
    if not root.is_dir():
        raise OfficeIdentityError('Native Office needs a private staging directory')
    deps = _load_dependencies(component)
    deps.initialize()
    try:
        before_processes = deps.processes()
        app = deps.dispatch(APPS[component][0])
        pid = _application_pid(app, component, deps)
        image = deps.process_image(pid)
        token = deps.identity(app)
        cls = NativeSheetComposer if component == 'spreadsheet' else NativeSlideComposer
        composer = cls.__new__(cls)
        composer.component, composer.pid, composer.executable = component, pid, image
        composer._app, composer._deps = app, deps
        composer._application_token = token
        composer._verify_application()
        composer._collection = getattr(app, APPS[component][3])
        composer._baseline = _snapshot(composer._collection, deps)
        composer._owns_app = pid not in before_processes and not composer._baseline
        composer._staging_dir = root
        composer._path = None
        composer._read_only = composer._visible = False
        composer._owns_doc = True
        composer._doc = composer._document_token = composer._ws = None
        composer._closed = False
        composer.cleanup_error = None
        composer._com_initialized = True
        # Bound process and complete baseline precede all mutation. The shared
        # app's security setting remains untouched; inputs are screened OOXML.
        if composer._owns_app:
            app.AutomationSecurity = 3
        if _snapshot(composer._collection, deps) != composer._baseline:
            raise OfficeIdentityError('Office documents changed before task creation')
        doc = (composer._collection.Add(Template=-4167) if component == 'spreadsheet'
               else composer._collection.Add(WithWindow=-1))
        document_token = deps.identity(doc)
        if document_token in [row[0] for row in composer._baseline]:
            raise OfficeIdentityError('Add returned a preexisting Office document')
        composer._doc, composer._document_token = doc, document_token
        composer._verify_document()
        return composer
    except BaseException:
        # An unverified return from Add/Open never grants permission to close
        # or Quit. The worker retains its private diagnostics for recovery.
        deps.uninitialize()
        raise
