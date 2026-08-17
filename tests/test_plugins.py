from __future__ import annotations

from pathlib import Path

import pytest

from skills.WPSComposer.scripts.plugins import run_plugins
from skills.WPSComposer.scripts.plugins import excalidraw


def test_unknown_requested_plugin_fails_closed():
    with pytest.raises(ValueError, match="Unknown plugin 'does-not-exist'"):
        run_plugins("# Report", ".", ["does-not-exist"])


def test_plugin_non_string_result_fails_closed(monkeypatch):
    monkeypatch.setitem(
        __import__(
            "skills.WPSComposer.scripts.plugins", fromlist=["_BUILTIN_PLUGINS"]
        )._BUILTIN_PLUGINS,
        "broken-result",
        lambda content, base_dir: None,
    )

    with pytest.raises(TypeError, match="must return Markdown text"):
        run_plugins("# Report", ".", ["broken-result"])


def test_excalidraw_plugin_never_overwrites_source_sidecar(monkeypatch, tmp_path):
    source = tmp_path / "diagram.excalidraw.md"
    source.write_text("{}", encoding="utf-8")
    sidecar = tmp_path / "diagram.excalidraw.png"
    sidecar.write_bytes(b"user-owned sidecar")
    rendered_paths = []

    def fake_render(source_path, output_path, width):
        rendered_paths.append(Path(output_path))
        Path(output_path).write_bytes(b"generated")
        return True

    monkeypatch.setattr(excalidraw, "_render_excalidraw_to_png", fake_render)

    result = excalidraw.excalidraw_plugin(
        "![[diagram.excalidraw.md]]", str(tmp_path)
    )

    assert sidecar.read_bytes() == b"user-owned sidecar"
    assert len(rendered_paths) == 1
    assert rendered_paths[0].parent != tmp_path
    assert str(rendered_paths[0]) in result


def test_requested_excalidraw_plugin_fails_closed_when_rendering_fails(
    monkeypatch, tmp_path
):
    source = tmp_path / "diagram.excalidraw.md"
    source.write_text("{}", encoding="utf-8")
    monkeypatch.setattr(
        excalidraw, "_render_excalidraw_to_png", lambda *args: False
    )

    with pytest.raises(RuntimeError, match="Failed to render Excalidraw"):
        run_plugins(
            "![[diagram.excalidraw.md]]", str(tmp_path), ["excalidraw"]
        )
