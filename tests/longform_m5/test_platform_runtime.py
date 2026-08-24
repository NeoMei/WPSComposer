from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import time

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.longform.executor import ExecutionOutcome, PaginationMap
from skills.WPSComposer.scripts.longform.quality import GenerationOutcome
from skills.WPSComposer.scripts.longform.relayout import RelayoutDirective
from skills.WPSComposer.scripts.longform import platform_runtime
from skills.WPSComposer.scripts.writer import WriterComposer


def test_relayout_derives_plan_and_only_sets_requested_heading_keep():
    build = build_longform_generation("# Report\n\n## Risk\n\nBody")
    heading = next(
        operation
        for operation in build.plan.operations
        if operation.op == "writer.add_heading"
    )
    derived = platform_runtime._apply_relayout(
        build.plan,
        (RelayoutDirective("keep-heading", heading.node_id, {}),),
    )
    original_heading = next(
        operation for operation in build.plan.operations if operation.op == "writer.add_heading"
    )
    derived_heading = next(
        operation for operation in derived.operations if operation.op == "writer.add_heading"
    )
    assert "keepWithNext" not in original_heading.args
    assert derived_heading.args["keepWithNext"] is True
    assert derived is not build.plan


def test_table_compression_is_a_closed_validated_m5_relayout():
    fixture = Path("tests/longform_m5/fixtures/wide_objects.md")
    build = build_longform_generation(
        fixture.read_text(encoding="utf-8"), base_dir=str(fixture.parent)
    )
    table = next(
        operation for operation in build.plan.operations
        if operation.op == "writer.add_semantic_table"
    )
    derived = platform_runtime._apply_relayout(
        build.plan,
        (RelayoutDirective("compress-table", table.node_id, {}),),
    )
    compressed = next(
        operation for operation in derived.operations
        if operation.op == "writer.add_semantic_table"
    )
    assert compressed.args["m5Relayout"] == "compress-table"
    assert compressed.args["allowRowSplit"] is True


def test_landscape_table_keeps_immediately_preceding_heading_and_continuous_exit():
    fixture = Path("tests/longform_m5/fixtures/wide_objects.md")
    build = build_longform_generation(
        fixture.read_text(encoding="utf-8"), base_dir=str(fixture.parent)
    )
    derived = platform_runtime._apply_relayout(build.plan, ())
    table = next(
        operation for operation in derived.operations
        if operation.op == "writer.add_semantic_table"
    )
    assert table.args["includePreviousHeading"] is True
    assert table.args["continuousExit"] is True


def test_blank_page_relayout_compacts_only_the_terminal_empty_paragraph():
    build = build_longform_generation("# Report\n\nBody")
    derived = platform_runtime._apply_relayout(
        build.plan,
        (RelayoutDirective("remove-unexpected-blank", None, {}),),
    )
    finalizer = next(
        operation for operation in derived.operations
        if operation.op == "writer.finalize_fields"
    )
    assert finalizer.args["compactTerminalParagraph"] is True


def test_vector_media_is_not_misclassified_as_zero_dpi_raster():
    fixture = Path("tests/longform_m5/fixtures/wide_objects.md")
    build = build_longform_generation(
        fixture.read_text(encoding="utf-8"), base_dir=str(fixture.parent)
    )
    _, policy = platform_runtime._quality_context(
        build,
        ExecutionOutcome("staged.docx", pagination_map=PaginationMap("M5-v1")),
    )
    assert policy.effective_dpi == ()


def test_windows_writer_detects_an_already_managed_landscape_section():
    class Sections:
        Count = 1

        def __call__(self, index):
            assert index == 1
            return SimpleNamespace(PageSetup=SimpleNamespace(Orientation=1))

    composer = SimpleNamespace(
        _doc=SimpleNamespace(
            Sections=Sections(), PageSetup=SimpleNamespace(Orientation=0)
        )
    )
    assert WriterComposer._current_section_is_landscape(composer) is True


def test_macos_pdf_export_waits_for_async_artifact_readiness(monkeypatch, tmp_path):
    captured = {}

    class Bridge:
        def issue(self, component, method, args):
            captured.update(args)
            return SimpleNamespace(id="cmd")

        def wait_result(self, command_id, timeout):
            return SimpleNamespace(ok=True, value={"path": captured["outputPath"]})

    adapter = platform_runtime.MacLongformAdapter.__new__(
        platform_runtime.MacLongformAdapter
    )
    adapter.executor = object()
    adapter.bridge = Bridge()
    adapter.runtime = SimpleNamespace(staging_dir=tmp_path)
    waited = []
    monkeypatch.setattr(
        platform_runtime,
        "_wait_for_pdf_artifact",
        lambda path, deadline: waited.append((Path(path), deadline)),
    )
    source = tmp_path / "source.docx"
    source.write_bytes(b"docx")
    deadline = time.monotonic() + 10
    result = adapter.export_pdf(source, deadline)
    assert result == Path(captured["outputPath"])
    assert waited == [(result, deadline)]


def test_generate_longform_always_closes_platform_adapter(monkeypatch, tmp_path):
    events = []

    class Adapter:
        def __init__(self, build):
            events.append("open")

        def close(self):
            events.append("close")

        def analyze(self, *args):
            raise AssertionError("stub lifecycle must not call analyze")

    monkeypatch.setattr(platform_runtime.sys, "platform", "darwin")
    monkeypatch.setattr(platform_runtime, "MacLongformAdapter", Adapter)
    monkeypatch.setattr(
        platform_runtime,
        "run_longform_lifecycle",
        lambda *args, **kwargs: GenerationOutcome(str(tmp_path / "result.docx")),
    )
    result = platform_runtime.generate_longform(
        object(),
        format_name="docx",
        output=tmp_path / "result.docx",
        timeout=10,
        overwrite=False,
    )
    assert result.path.endswith("result.docx")
    assert events == ["open", "close"]
