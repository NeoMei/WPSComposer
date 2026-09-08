"""Explicit engine selection without launching or modifying Office applications."""
from __future__ import annotations

from contextlib import contextmanager
from contextvars import ContextVar
import math
import os
from pathlib import Path
import re
import sys

ENGINES = ('wps', 'msoffice', 'auto')
_COMPONENTS = ('writer', 'spreadsheet', 'presentation')
_OFFICE_PROGIDS = {'Word.Application', 'Excel.Application', 'PowerPoint.Application'}
_COM_ENGINE = ContextVar('wpscomposer_com_engine', default=None)


class EngineUnavailableError(RuntimeError):
    """Engine/component cannot be selected before a native task begins."""
    code = 'BACKEND_UNAVAILABLE'


def validate_engine(engine: str) -> str:
    if not isinstance(engine, str):
        raise TypeError('engine must be wps, msoffice or auto')
    if engine not in ENGINES:
        raise ValueError('engine must be wps, msoffice or auto')
    return engine


def validate_timeout(timeout: float) -> float:
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)):
        raise TypeError('timeout must be a positive finite number')
    if not math.isfinite(timeout) or timeout <= 0:
        raise ValueError('timeout must be a positive finite number')
    return float(timeout)


def _registered_executable(progids, allowed_names):
    # Query the current Python bitness, the same view COM dispatch will use.
    import winreg
    for progid in progids:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r'\CLSID') as key:
                clsid = winreg.QueryValueEx(key, None)[0]
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + clsid + r'\LocalServer32') as key:
                command = winreg.QueryValueEx(key, None)[0]
            command = os.path.expandvars(command).strip()
            match = re.match(r'^"([^"]+\.exe)"|^(.+?\.exe)(?:\s|$)', command, re.I)
            if match:
                path = Path(match.group(1) or match.group(2))
                if path.name.lower() in allowed_names and path.is_file():
                    return str(path)
        except OSError:
            continue
    return None


def engine_executable(engine: str, component: str = 'writer'):
    """Return installed native executable/app, or None; never start an app."""
    validate_engine(engine)
    if component not in _COMPONENTS or engine == 'auto':
        raise ValueError('Detection requires an explicit engine and valid component')
    if engine == 'msoffice' and component != 'writer':
        return None
    if sys.platform == 'darwin':
        # Match locations accepted by the native execution adapters.
        name = 'wpsoffice.app' if engine == 'wps' else 'Microsoft Word.app'
        candidate = Path('/Applications') / name
        return str(candidate) if candidate.is_dir() else None
    if sys.platform == 'win32':
        if engine == 'msoffice':
            return _registered_executable(('Word.Application',), {'winword.exe'})
        progids = {'writer': ('KWps.Application', 'Wps.Application'),
                   'spreadsheet': ('Ket.Application',),
                   'presentation': ('KWpp.Application', 'Wpp.Application')}[component]
        return _registered_executable(progids, {'wps.exe', 'et.exe', 'wpp.exe'})
    return None


def resolve_engine(engine: str, component: str) -> str:
    validate_engine(engine)
    if component not in _COMPONENTS:
        raise ValueError('Unknown Office component: ' + str(component))
    if engine == 'msoffice' and component != 'writer':
        raise EngineUnavailableError('MS Office supports writer DOCX/PDF in this release; use engine="wps" for other components')
    if engine != 'auto':
        return engine
    for candidate in ('wps', 'msoffice'):
        if candidate == 'msoffice' and component != 'writer':
            continue
        if engine_executable(candidate, component):
            return candidate
    raise EngineUnavailableError('No installed native engine for ' + component)


@contextmanager
def com_engine(engine: str):
    token = _COM_ENGINE.set(validate_engine(engine))
    try:
        yield
    finally:
        _COM_ENGINE.reset(token)


def com_progids(progids):
    """Restrict dispatch for a pinned public task; direct legacy composers retain their API."""
    values = tuple(progids)
    if _COM_ENGINE.get() == 'wps':
        values = tuple(p for p in values if p not in _OFFICE_PROGIDS)
        if not values:
            raise EngineUnavailableError('No WPS ProgID for requested component')
    return values
