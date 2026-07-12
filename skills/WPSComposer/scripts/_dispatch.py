"""COM dispatch layer — ProgID chains, format constants, WPS path search.

Shared infrastructure for all WPS / Office COM composers.
"""

from __future__ import annotations

import os
import platform

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
FMT_PPT = 2                         # .ppt (97-2003)
FMT_PPSX = 36                       # .ppsx
FMT_PPTM = 25                       # macro-enabled presentation
FMT_POTX = 26                       # template
FMT_POTM = 27                       # macro-enabled template
FMT_PPSM = 28                       # macro-enabled show
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
    import win32com.client as win32
    import pythoncom
    pythoncom.CoInitialize()
    last = None
    for pid in progids:
        try:
            try:
                return win32.DispatchEx(pid)
            except Exception:
                return win32.Dispatch(pid)
        except Exception as e:
            last = e
    raise WPSUnavailable(f"No COM host for {progids}: {last}")


def _safe_quit(app):
    """Quit a COM app, swallowing any errors."""
    try:
        app.Quit()
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
