from __future__ import annotations

import sys
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts import _base, _dispatch


class _App:
    def __init__(self):
        self.quit_calls = 0
        self.Visible = -1
        self.DisplayAlerts = -1

    def Quit(self):
        self.quit_calls += 1


class _Composer(_base.BaseComposer):
    _progids = ("Wps.Application",)

    @staticmethod
    def _create_doc(app):
        return SimpleNamespace(Close=lambda *args: None)

    @staticmethod
    def _open_document(app, path, read_only=False):
        return SimpleNamespace(Close=lambda *args: None)

    @staticmethod
    def _active_document(app):
        return getattr(app, "active_document", None)


def _install_fake_com(monkeypatch, *, dispatch_ex, dispatch=None, active=None):
    calls = []

    pythoncom = SimpleNamespace(
        CoInitialize=lambda: calls.append("init"),
        CoUninitialize=lambda: calls.append("uninit"),
    )

    def get_active(progid):
        calls.append(("active", progid))
        if isinstance(active, BaseException):
            raise active
        return active

    client = SimpleNamespace(
        DispatchEx=dispatch_ex,
        Dispatch=dispatch,
        GetActiveObject=get_active,
    )
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(_dispatch.platform, "system", lambda: "Windows")
    return calls


def test_dispatch_ex_application_is_owned_and_com_apartment_is_balanced(monkeypatch):
    app = _App()
    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=lambda progid: app,
        dispatch=lambda progid: pytest.fail("Dispatch fallback should not run"),
    )

    with _Composer() as composer:
        assert composer.app is app

    assert app.quit_calls == 1
    assert calls == ["init", "uninit"]


def test_dispatch_fallback_is_shared_and_is_never_quit(monkeypatch):
    app = _App()

    def unavailable_factory(progid):
        raise RuntimeError("no local-server factory")

    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=unavailable_factory,
        dispatch=lambda progid: app,
    )

    with _Composer() as composer:
        assert composer.app is app

    assert app.quit_calls == 0
    assert app.Visible == -1
    assert app.DisplayAlerts == -1
    assert calls == ["init", "uninit"]


def test_dispatch_failure_uninitializes_com_apartment(monkeypatch):
    def unavailable(progid):
        raise RuntimeError(progid)

    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=unavailable,
        dispatch=unavailable,
    )

    with pytest.raises(_dispatch.WPSUnavailable):
        _Composer().__enter__()

    assert calls == ["init", "uninit"]


@pytest.mark.parametrize("owned", [True, False])
def test_document_open_failure_releases_partial_application_and_apartment(
    monkeypatch, owned
):
    app = _App()
    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=(lambda progid: app) if owned else (
            lambda progid: (_ for _ in ()).throw(RuntimeError("shared"))
        ),
        dispatch=lambda progid: app,
    )

    class FailingComposer(_Composer):
        @staticmethod
        def _create_doc(app):
            raise RuntimeError("open failed")

    with pytest.raises(RuntimeError, match="open failed"):
        FailingComposer().__enter__()

    assert app.quit_calls == (1 if owned else 0)
    assert calls == ["init", "uninit"]


def test_attach_active_balances_apartment_without_closing_user_app(monkeypatch):
    app = _App()
    app.active_document = object()
    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=lambda progid: pytest.fail("unused"),
        dispatch=lambda progid: pytest.fail("unused"),
        active=app,
    )

    composer = _Composer.attach_active()
    composer.close()
    composer.close()

    assert app.quit_calls == 0
    assert calls == ["init", ("active", "Wps.Application"), "uninit"]


def test_attach_active_failure_uninitializes_apartment(monkeypatch):
    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=lambda progid: pytest.fail("unused"),
        dispatch=lambda progid: pytest.fail("unused"),
        active=RuntimeError("not running"),
    )

    with pytest.raises(_dispatch.WPSUnavailable):
        _Composer.attach_active()

    assert calls == ["init", ("active", "Wps.Application"), "uninit"]


def test_interrupted_document_open_still_releases_owned_app_and_apartment(monkeypatch):
    app = _App()
    calls = _install_fake_com(
        monkeypatch,
        dispatch_ex=lambda progid: app,
        dispatch=lambda progid: pytest.fail("unused"),
    )

    class InterruptedComposer(_Composer):
        @staticmethod
        def _create_doc(app):
            raise KeyboardInterrupt()

    with pytest.raises(KeyboardInterrupt):
        InterruptedComposer().__enter__()

    assert app.quit_calls == 1
    assert calls == ["init", "uninit"]


def test_attached_save_copy_fails_before_rebinding_without_savecopyas(tmp_path):
    save_as_calls = []

    class AttachedDocument:
        FullName = r"C:\\user\\original.docx"
        SaveFormat = 12

        def SaveCopyAs(self, path):
            raise AttributeError("SaveCopyAs is unavailable")

        def SaveAs(self, *args):
            save_as_calls.append(args)

    composer = _Composer()
    composer._doc = AttachedDocument()
    composer._owns_doc = False

    with pytest.raises(RuntimeError, match="non-rebinding copy primitive"):
        composer.save_copy(tmp_path / "copy.docx")

    assert save_as_calls == []
    assert composer.doc.FullName == r"C:\\user\\original.docx"
