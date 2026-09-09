"""COM dispatch layer — ProgID chains, format constants, WPS path search.

Shared infrastructure for all WPS / Office COM composers.
"""

from __future__ import annotations

import atexit
import os
import platform
import threading
import time
from dataclasses import dataclass

# ---------------------------------------------------------------------------
# ProgID fallback chains
# ---------------------------------------------------------------------------

WRITER_PROGIDS = ("KWps.Application", "Wps.Application", "Word.Application")
SHEET_PROGIDS = ("Ket.Application", "Excel.Application")
SLIDE_PROGIDS = ("KWpp.Application", "Wpp.Application", "PowerPoint.Application")

# ---------------------------------------------------------------------------
# Export format constants (WPS / Office compatible)
# ---------------------------------------------------------------------------

# Writer formats
FMT_DOCX = 12
FMT_PDF_FROM_DOC = 17
FMT_DOC = 0                         # .doc (97-2003)
FMT_DOCM = 13                       # macro-enabled document
FMT_DOTX = 14                       # template
FMT_DOTM = 15                       # macro-enabled template
FMT_TXT = 2                         # .txt
FMT_HTML = 8                        # .html
FMT_MHTML = 9                       # .mht/.mhtml
FMT_RTF = 6                         # .rtf
FMT_XML = 11                        # .xml
FMT_ODT = 23                        # .odt
FMT_XPS = 18                        # .xps

# Sheet formats
FMT_XLSX = 51
FMT_PDF_FROM_XLS = 0
FMT_XLS = -4143                     # .xls (97-2003)
FMT_XLSM = 52                       # macro-enabled workbook
FMT_XLSB = 50                       # binary workbook
FMT_XLTX = 54                       # template
FMT_XLTM = 53                       # macro-enabled template
FMT_CSV = 62                        # .csv (UTF-8)
FMT_TSV = -4158                     # tab-delimited text
FMT_ODS = 60                        # OpenDocument spreadsheet

# Slide formats
FMT_PPTX = 24
FMT_PDF_FROM_PPT = 32
FMT_PPT = 1                         # .ppt (97-2003, ppSaveAsPresentation)
FMT_PPSX = 28                       # .ppsx (ppSaveAsOpenXMLShow)
FMT_PPTM = 25                       # macro-enabled presentation
FMT_POTX = 26                       # template
FMT_POTM = 27                       # macro-enabled template
FMT_PPSM = 29                       # macro-enabled show
FMT_ODP = 35                        # OpenDocument presentation

# ---------------------------------------------------------------------------
# WPS installation search paths
# ---------------------------------------------------------------------------

WPS_SEARCH_PATHS = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 r"Kingsoft\WPS Office"),
    os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                 r"Kingsoft\WPS Office"),
    os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                 r"Kingsoft\WPS Office"),
    r"C:\Program Files (x86)\Kingsoft\WPS Office",
    r"C:\Program Files\Kingsoft\WPS Office",
]


class WPSUnavailable(RuntimeError):
    """Raised when no COM host (WPS or Office) is available."""
    pass


@dataclass(frozen=True)
class DispatchResult:
    """A COM application together with its application ownership."""

    app: object
    owns_app: bool


def _require():
    """Verify Windows + pywin32 are available."""
    if platform.system() != "Windows":
        raise WPSUnavailable("WPS COM requires Windows.")
    try:
        import win32com.client  # noqa: F401
        import pythoncom        # noqa: F401
    except Exception as exc:
        raise WPSUnavailable(f"pywin32 not installed: {exc}")


def _dispatch(progids):
    """Try each ProgID in order; return a dedicated automation instance.

    ``DispatchEx`` avoids reusing and later quitting the user's interactive WPS
    process.  Some WPS builds do not expose a local-server factory, so a normal
    ``Dispatch`` remains the compatibility fallback.
    """
    from .office_engines import com_progids
    progids = com_progids(progids)
    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()
    try:
        last = None
        for pid in progids:
            try:
                return DispatchResult(win32.DispatchEx(pid), True)
            except Exception:
                try:
                    return DispatchResult(win32.Dispatch(pid), False)
                except Exception as exc:
                    last = exc
        raise WPSUnavailable(f"No COM host for {progids}: {last}")
    except BaseException:
        pythoncom.CoUninitialize()
        raise


def _safe_quit(app):
    """Quit a COM app and wait (bounded) for the host to actually exit.

    WPS Quit is asynchronous; a following DispatchEx in the same process can
    otherwise receive a proxy into the still-quitting instance and die with a
    mid-run RPC error.
    """
    try:
        app.Quit()
    except BaseException:
        pass
    deadline = time.monotonic() + 3.0
    while time.monotonic() < deadline:
        try:
            app.Version  # any cheap property round-trip
        except BaseException:
            return
        time.sleep(0.2)


# ---------------------------------------------------------------------------
# Pooled suite applications
#
# On suite builds (e.g. WPS Office personal zh-CN) the Writer, Presentation,
# and Spreadsheet automation servers all live inside ONE wps.exe host
# process.  Quitting any suite app mid-process tears down the shared host
# and kills every other pooled instance with "object not connected to
# server" / mid-run RPC errors.  Pooled suite apps are therefore NEVER quit
# during the process lifetime; they are quit once at interpreter exit.
# ---------------------------------------------------------------------------

_POOLED_SUITE_APPS: dict[tuple, object] = {}


def _pool_probe(app) -> None:
    """Liveness/readiness probe: touch the primary collection of the app."""
    for collection in ("Documents", "Presentations", "Workbooks"):
        try:
            int(getattr(app, collection).Count)
            return
        except Exception:
            continue
    raise WPSUnavailable("pooled suite app is not responsive")


def pooled_suite_app(progids):
    """Return the pooled dedicated app for this ProgID chain and thread.

    The pool never falls back to a shared ``Dispatch``: pooled callers are
    headless generation paths that must own their instance.
    """
    import pythoncom  # pywin32
    import win32com.client as win32

    from .office_engines import com_progids
    progids = com_progids(progids)
    key = (tuple(progids), threading.get_ident())
    app = _POOLED_SUITE_APPS.get(key)
    if app is not None:
        try:
            _pool_probe(app)
            return app
        except Exception:
            del _POOLED_SUITE_APPS[key]

    # ponytail: this CoInitialize is intentionally unbalanced — the pooled
    # app is apartment-affine and must outlive every composer on this
    # thread; pythoncom cleans up at thread/process teardown anyway.
    pythoncom.CoInitialize()
    app = None
    last: Exception | None = None
    for pid in progids:
        try:
            app = win32.DispatchEx(pid)
            _pool_probe(app)
            break
        except Exception as exc:
            app = None
            last = exc
    if app is None:
        pythoncom.CoUninitialize()
        raise WPSUnavailable(f"No dedicated COM host for {progids}: {last}")
    try:
        app.Visible = 0
    except Exception:
        pass
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    _POOLED_SUITE_APPS[key] = app
    return app


@atexit.register
def _quit_pooled_suite_apps() -> None:  # pragma: no cover - teardown
    while _POOLED_SUITE_APPS:
        _key, app = _POOLED_SUITE_APPS.popitem()
        try:
            _safe_quit(app)
        except Exception:
            pass


def _abs(path):
    """Normalise a path for COM: absolute, backslashes."""
    return os.path.abspath(path).replace("/", "\\")


def find_wps_executable():
    """Locate the latest wps.exe on this machine.

    Searches known WPS installation directories, then falls back to PATH.
    Returns the path to wps.exe or "wps.exe" as a last resort.
    """
    candidates = []
    for base in WPS_SEARCH_PATHS:
        if not os.path.isdir(base):
            continue
        for item in os.listdir(base):
            item_path = os.path.join(base, item)
            if not os.path.isdir(item_path):
                continue
            wps_exe = os.path.join(item_path, "office6", "wps.exe")
            if os.path.isfile(wps_exe):
                candidates.append(wps_exe)
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0]
    import shutil
    return shutil.which("wps.exe") or "wps.exe"
