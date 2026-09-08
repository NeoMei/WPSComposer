"""Failure-path ownership and deadline regressions without native Word."""
from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path
from types import ModuleType, SimpleNamespace

import pytest


@pytest.fixture
def sentinel_runner(monkeypatch):
    scripts = Path(__file__).resolve().parents[1] / "fixtures" / "msoffice_spike"
    for name in ("mac_word", "windows_word", "windows_unsaved_sentinel"):
        spec = importlib.util.spec_from_file_location(name, scripts / (name + ".py"))
        module = importlib.util.module_from_spec(spec)
        monkeypatch.setitem(sys.modules, name, module)
        spec.loader.exec_module(module)
    return sys.modules["windows_unsaved_sentinel"]


def install_word(monkeypatch, runner, failure=None, change=None):
    """Replace only the external COM boundary; run the real runner and reports."""
    class Document:
        Name = "Document1"
        FullName = "Document1"
        Saved = True
        closed = False

        def __init__(self):
            self.text = "\r"
            self.Content = self

        @property
        def Text(self):
            return self.text

        @Text.setter
        def Text(self, value):
            if failure == "marker_write":
                raise RuntimeError("marker write failed")
            self.text = value + "\r"
            self.Saved = False
            if failure == "marker_written_then_error":
                raise RuntimeError("marker setter reported an error after writing")

        @property
        def Windows(self):
            if failure == "window":
                if change is not None:
                    change(self)
                raise RuntimeError("HWND unavailable")
            return SimpleNamespace(Item=lambda index: SimpleNamespace(Hwnd=123))

        def Close(self, SaveChanges):
            assert SaveChanges == 0
            self.closed = True
            app.Documents.Count = 0

    doc = Document()

    def add():
        app.Documents.Count = 1
        return doc

    app = SimpleNamespace(Name="Microsoft Word", Version="16", Build="test", Path="C:/Office",
                          Documents=SimpleNamespace(Count=0, Add=add, Item=lambda index: doc),
                          _oleobj_=SimpleNamespace(QueryInterface=lambda iid: "same-application"))
    win32com = ModuleType("win32com")
    client = ModuleType("win32com.client")
    client.GetActiveObject = lambda name: app
    win32com.client = client
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setitem(sys.modules, "pythoncom", SimpleNamespace(CoInitialize=lambda: None, IID_IUnknown="IUnknown"))
    monkeypatch.setitem(sys.modules, "win32process", SimpleNamespace(GetWindowThreadProcessId=lambda hwnd: (1, 42)))
    monkeypatch.setattr(runner, "word_processes", lambda: {42: "C:/Office/WINWORD.EXE"})
    return doc


def arguments(monkeypatch, tmp_path, extra=()):
    evidence = tmp_path / "evidence"
    output = tmp_path / "native"
    monkeypatch.setattr(sys, "argv", ["sentinel", "--evidence-dir", str(evidence),
                                      "--output-dir", str(output), *extra])
    return evidence, output


@pytest.mark.parametrize("failure", ["window", "marker_write", "marker_written_then_error"])
def test_closes_unchanged_new_document_after_early_setup_failure(sentinel_runner, monkeypatch, tmp_path, failure):
    doc = install_word(monkeypatch, sentinel_runner, failure=failure)
    evidence, _ = arguments(monkeypatch, tmp_path)
    assert sentinel_runner.main() == 1
    assert doc.closed
    report = json.loads((evidence / "sentinel-result.json").read_text())
    assert report["sentinel_close_without_saving"] is True
    assert report["overall"] == "failed"


@pytest.mark.parametrize("attribute,value", [("text", "user text\r"), ("Name", "changed"),
                                             ("FullName", "C:/user.docx"), ("Saved", False)])
def test_refuses_changed_initial_document(sentinel_runner, monkeypatch, tmp_path, attribute, value):
    doc = install_word(monkeypatch, sentinel_runner, failure="window",
                       change=lambda current: setattr(current, attribute, value))
    evidence, _ = arguments(monkeypatch, tmp_path)
    assert sentinel_runner.main() == 1
    assert not doc.closed
    assert json.loads((evidence / "sentinel-result.json").read_text())["sentinel_close_without_saving"] is False


@pytest.mark.parametrize("attribute,value", [("text", "user text\r"), ("Name", "changed"),
                                             ("FullName", "C:/user.docx"), ("Saved", True)])
def test_refuses_changed_marked_document(sentinel_runner, monkeypatch, tmp_path, attribute, value):
    doc = install_word(monkeypatch, sentinel_runner)
    evidence, _ = arguments(monkeypatch, tmp_path)

    def run(command, **kwargs):
        setattr(doc, attribute, value)
        raise RuntimeError("helper failed after sentinel changed")

    monkeypatch.setattr(sentinel_runner.subprocess, "run", run)
    assert sentinel_runner.main() == 1
    assert not doc.closed
    assert json.loads((evidence / "sentinel-result.json").read_text())["sentinel_close_without_saving"] is False


@pytest.mark.parametrize("custom_timeout", [None, "0.25"])
def test_timeout_retains_partial_evidence_and_closes_only_sentinel(sentinel_runner, monkeypatch, tmp_path, custom_timeout):
    doc = install_word(monkeypatch, sentinel_runner)
    evidence, output = arguments(monkeypatch, tmp_path,
                                 () if custom_timeout is None else ("--timeout-seconds", custom_timeout))

    def run(command, **kwargs):
        timeout = kwargs.get("timeout")
        assert timeout == (660 if custom_timeout is None else 0.25)
        output.mkdir()
        (output / "result.json").write_text('{"overall":"in_progress"}')
        (output / "probe.docx").write_bytes(b"partial native evidence")
        raise subprocess.TimeoutExpired(command, timeout, output=b"partial stdout", stderr=b"partial stderr")

    monkeypatch.setattr(sentinel_runner.subprocess, "run", run)
    assert sentinel_runner.main() == 1
    report = json.loads((evidence / "sentinel-result.json").read_text())
    assert report["runner_timed_out"] is True
    assert "unverified" in report["runner_cleanup_warning"]
    assert (evidence / "runner.stdout.log").read_bytes() == b"partial stdout"
    assert (evidence / "runner.stderr.log").read_bytes() == b"partial stderr"
    assert (output / "probe.docx").read_bytes() == b"partial native evidence"
    assert json.loads((output / "result.json").read_text()) == {"overall": "in_progress"}
    assert doc.closed and report["application_quit_attempted"] is False


@pytest.mark.parametrize("value", ["0", "-1", "nan", "inf", "-inf"])
def test_rejects_invalid_timeout_before_creating_directories(sentinel_runner, monkeypatch, tmp_path, value):
    evidence, output = arguments(monkeypatch, tmp_path, ("--timeout-seconds=" + value,))
    with pytest.raises(SystemExit) as exc:
        sentinel_runner.main()
    assert exc.value.code == 2
    assert not evidence.exists() and not output.exists()


@pytest.mark.parametrize("relationship", ["equal", "output_parent", "evidence_parent"])
def test_rejects_overlapping_paths_before_creating_directories(sentinel_runner, monkeypatch, tmp_path, relationship):
    evidence = tmp_path / "fresh"
    output = evidence if relationship == "equal" else evidence / "child"
    if relationship == "output_parent":
        evidence, output = output, evidence
    monkeypatch.setattr(sys, "argv", ["sentinel", "--evidence-dir", str(evidence), "--output-dir", str(output)])
    with pytest.raises(SystemExit) as exc:
        sentinel_runner.main()
    assert exc.value.code == 2
    assert not evidence.exists() and not output.exists()
