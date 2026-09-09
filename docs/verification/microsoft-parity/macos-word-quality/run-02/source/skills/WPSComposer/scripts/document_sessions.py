"""Lazy dispatch to engine-bound Microsoft document sessions.

Once returned, a session is the sole authority for its document. Patch calls
operate on that session; they never discover or switch an engine again.
"""
from __future__ import annotations

import importlib
import sys

from .office_engines import EngineUnavailableError


def _session_class(kind):
    names = {'writer': ('word', 'Word'), 'sheet': ('excel', 'Excel'),
             'slide': ('powerpoint', 'PowerPoint')}
    if kind not in names:
        raise ValueError('kind must be writer, sheet, or slide')
    if sys.platform == 'win32':
        module_name, name = '.msoffice.windows_session_proxy', 'Proxy' + names[kind][1] + 'Session'
    elif sys.platform == 'darwin':
        module_name, name = '.msoffice.macos_' + names[kind][0] + '_session', 'Mac' + names[kind][1] + 'Session'
    else:
        raise EngineUnavailableError('Native Microsoft document sessions require Windows or macOS')
    try:
        module = importlib.import_module(module_name, __package__)
    except ModuleNotFoundError as exc:
        expected = __package__ + module_name
        if exc.name != expected:
            raise
        raise EngineUnavailableError('Native Microsoft session adapter is not implemented for ' + kind) from None
    return getattr(module, name)


def open_session(path, *, kind, read_only=False, visible=False):
    return _session_class(kind).open_document(path, read_only=read_only, visible=visible)


def new_session(*, kind, visible=False):
    return _session_class(kind).new_document(visible=visible)


def attach_session(*, kind=None):
    if kind is None:
        raise ValueError('Specify kind when attaching to a Microsoft application')
    return _session_class(kind).attach_active()
