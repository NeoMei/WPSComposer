"""Exercise cleanup decisions around native Word identity failures without COM."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace

import pytest


@pytest.mark.parametrize("failure", ["add", "name", "window", "pid_lookup", "pid_mismatch", "after_identity"])
def test_quit_requires_document_process_identity(monkeypatch, tmp_path, failure):
    scripts = Path(__file__).resolve().parents[1] / "fixtures" / "msoffice_spike"
    for name in ("mac_word", "windows_word"):
        spec = importlib.util.spec_from_file_location(name, scripts / (name + ".py"))
        module = importlib.util.module_from_spec(spec)
        monkeypatch.setitem(sys.modules, name, module)
        spec.loader.exec_module(module)
    runner = sys.modules["windows_word"]
    calls = []

    class Document:
        @property
        def Name(self):
            if failure == "name":
                raise RuntimeError("name unavailable")
            return "owned test document"

        @property
        def Windows(self):
            if failure == "window":
                raise RuntimeError("window handle unavailable")
            return SimpleNamespace(Item=lambda index: SimpleNamespace(Hwnd=123))

        def Close(self, SaveChanges):
            assert SaveChanges == 0
            calls.append("close_owned_document")
            app.Documents.Count = 0

        def __setattr__(self, name, value):
            if name == "RemovePersonalInformation":
                raise RuntimeError("failure after identity verification")
            object.__setattr__(self, name, value)

    class Application:
        Version = "16.0"
        Build = "test"

        def __init__(self):
            self.Documents = SimpleNamespace(Count=0, Add=self.add)

        def add(self):
            if failure == "add":
                raise RuntimeError("document creation failed")
            self.Documents.Count = 1
            return Document()

        def Quit(self, SaveChanges):
            assert SaveChanges == 0 and self.Documents.Count == 0
            calls.append("quit_application")

        def __setattr__(self, name, value):
            if name == "Visible":
                calls.append("change_visibility")
            object.__setattr__(self, name, value)

    class NoActiveWord(Exception):
        hresult = -2147221021

    def get_active(_):
        raise NoActiveWord()

    def window_pid(hwnd):
        assert hwnd == 123
        if failure == "pid_lookup":
            raise RuntimeError("PID lookup failed")
        return 1, 999 if failure == "pid_mismatch" else 42

    app = Application()
    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.GetActiveObject = get_active
    client.DispatchEx = lambda progid: app
    win32com.client = client
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setitem(sys.modules, "pythoncom", SimpleNamespace(CoInitialize=lambda: None))
    monkeypatch.setitem(sys.modules, "win32process", SimpleNamespace(GetWindowThreadProcessId=window_pid))
    monkeypatch.setattr(runner, "word_processes", lambda: {})
    monkeypatch.setattr(runner, "application_identity", lambda app, before: {"pid": 42})
    monkeypatch.setattr(runner.platform, "platform", lambda: "Windows test double")
    monkeypatch.setattr(sys, "platform", "win32")
    output = tmp_path / "fresh-run"
    monkeypatch.setattr(sys, "argv", ["windows_word.py", "--output-dir", str(output)])

    assert runner.main() == 1
    assert json.loads((output / "result.json").read_text())["overall"] == "failed"
    assert ("close_owned_document" in calls) == (failure != "add")
    identity_verified = failure == "after_identity"
    assert ("quit_application" in calls) == identity_verified
    assert ("change_visibility" in calls) == identity_verified
