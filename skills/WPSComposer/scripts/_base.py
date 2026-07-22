"""Enhanced base composer — shared COM lifecycle, save, and export logic.

WriterComposer, SheetComposer, and SlideComposer inherit from this.
"""

from __future__ import annotations

import os

from ._dispatch import _require, _dispatch, _safe_quit, _abs


class BaseComposer:
    """Common lifecycle for WPS COM composers.

    Subclasses must set:
        - ``_progids``   : tuple of ProgID strings to try.
        - ``_doc_type``  : "writer" / "calc" / "impress".
        - ``_native_fmt``: format constant for the native save (e.g. FMT_DOCX).
        - ``_pdf_fmt``   : format constant for PDF export.
        - ``_create_doc(app)`` : classmethod or staticmethod returning a new
                                 COM document object.
    """

    _progids = ()
    _doc_type = ""
    _native_fmt = 0
    _pdf_fmt = 0
    _formats_by_extension = {}

    def __init__(self, path=None, read_only=False, visible=False):
        """Create a composer for a new or existing document.

        ``path`` is optional.  When supplied, entering the context opens the
        existing document instead of creating an empty one.  Use
        :meth:`attach_active` when the agent should operate on the document
        currently open in the user's WPS/Office window.
        """
        self._app = None
        self._doc = None
        self._path = _abs(path) if path else None
        self._read_only = bool(read_only)
        self._visible = bool(visible)
        self._owns_app = False
        self._owns_doc = False

    # ---- subclass hooks ----

    @staticmethod
    def _create_doc(app):
        raise NotImplementedError("subclass must implement _create_doc(app)")

    @staticmethod
    def _open_document(app, path, read_only=False):
        raise NotImplementedError("subclass must implement _open_document(app, path)")

    @staticmethod
    def _active_document(app):
        raise NotImplementedError("subclass must implement _active_document(app)")

    # ---- COM lifecycle ----

    def _open(self):
        _require()
        self._app = _dispatch(self._progids)
        self._owns_app = True
        try:
            self._app.Visible = -1 if self._visible else 0
        except Exception:
            pass
        try:
            self._app.DisplayAlerts = 0
        except Exception:
            pass
        return self._app

    def __enter__(self):
        if self._doc is not None:
            return self
        app = self._open()
        if self._path:
            self._doc = self._open_document(app, self._path, self._read_only)
        else:
            self._doc = self._create_doc(app)
        self._owns_doc = True
        return self

    def __exit__(self, *exc):
        self.close(save_changes=False)

    @classmethod
    def open_document(cls, path, read_only=False, visible=False):
        """Return a context-manageable composer for an existing file."""
        return cls(path=path, read_only=read_only, visible=visible)

    @classmethod
    def attach_active(cls):
        """Attach to the user's active WPS/Office document without owning it.

        Closing the returned composer never closes the document or quits the
        application.  This is the preferred entry point for conversational,
        fine-grained edits to the document currently on screen.
        """
        _require()
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        last = None
        app = None
        for progid in cls._progids:
            try:
                app = win32.GetActiveObject(progid)
                break
            except Exception as exc:
                last = exc
        if app is None:
            from ._dispatch import WPSUnavailable
            raise WPSUnavailable(
                f"No running WPS/Office application for {cls._progids}: {last}"
            )

        obj = cls()
        obj._app = app
        obj._doc = cls._active_document(app)
        if obj._doc is None:
            from ._dispatch import WPSUnavailable
            raise WPSUnavailable("The application is running but has no active document.")
        obj._owns_app = False
        obj._owns_doc = False
        return obj

    def close(self, save_changes=False):
        """Release owned COM objects; attached user documents stay open."""
        if self._owns_doc and self._doc is not None:
            try:
                try:
                    self._doc.Close(bool(save_changes))
                except TypeError:
                    # PowerPoint Presentation.Close() takes no argument
                    self._doc.Close()
            except Exception:
                pass
        if self._owns_app:
            _safe_quit(self._app)
        self._doc = None
        self._app = None

    # ---- shared I/O ----

    def save(self, path, fmt=None):
        """Save the document in *fmt* (defaults to native format).

        Returns the absolute path.
        """
        p = _abs(path)
        if fmt is not None:
            use_fmt = fmt
        else:
            ext = os.path.splitext(p)[1].lower()
            use_fmt = self._formats_by_extension.get(ext)
            if use_fmt is None and ext:
                # Let WPS choose its proprietary/host-specific format from the
                # extension (for example .wps/.et/.dps).
                self._doc.SaveAs(p)
                return p
            if use_fmt is None:
                use_fmt = self._native_fmt
        self._doc.SaveAs(p, use_fmt)
        return p

    def save_current(self):
        """Save an opened/attached document in its current format and path."""
        self._doc.Save()
        return _abs(self._doc.FullName)

    def export_pdf(self, path):
        """Export to PDF.  Returns the absolute path."""
        p = _abs(path)
        self._doc.SaveAs(p, self._pdf_fmt)
        return p

    def save_copy(self, path, fmt=None):
        """Save to another path while keeping the session available."""
        return self.save(path, fmt)

    # ---- properties ----

    @property
    def app(self):
        return self._app

    @property
    def doc(self):
        return self._doc
