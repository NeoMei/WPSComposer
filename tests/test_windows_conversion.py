from __future__ import annotations

from pathlib import Path

import pytest

from skills.WPSComposer.scripts.conversion import ConversionRequest
from skills.WPSComposer.scripts.windows_conversion import convert_windows


class FakeDocument:
    def __init__(self, calls: list):
        self.calls = calls

    @staticmethod
    def _write_pdf(path: str) -> None:
        Path(path).write_bytes(b"%PDF-1.7\n" + b"x" * 2048)

    def ExportAsFixedFormat(self, *args):
        self.calls.append(("ExportAsFixedFormat", args))
        path = args[0] if isinstance(args[0], str) else args[1]
        self._write_pdf(path)

    def SaveAs(self, *args):
        self.calls.append(("SaveAs", args))
        self._write_pdf(args[0])


class FakeComposer:
    def __init__(self, calls: list, source: str, read_only: bool, visible: bool):
        self.calls = calls
        self.doc = FakeDocument(calls)
        self.calls.append(("open_document", source, read_only, visible))

    def __enter__(self):
        self.calls.append(("enter",))
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.calls.append(("close", False))


def _composer_class(calls: list):
    class ComposerClass:
        @classmethod
        def open_document(cls, source, read_only=False, visible=False):
            return FakeComposer(calls, source, read_only, visible)

    return ComposerClass


def _request(tmp_path: Path, suffix: str, component: str) -> ConversionRequest:
    source = tmp_path / f"input.{suffix}"
    source.write_bytes(b"source")
    return ConversionRequest(
        source=source.resolve(),
        output=(tmp_path / f"{component}.pdf").resolve(),
        component=component,
        overwrite=False,
    )


def test_writer_uses_document_export_as_fixed_format(tmp_path: Path):
    calls = []
    request = _request(tmp_path, "docx", "writer")

    result = convert_windows(
        request, composer_classes={"writer": _composer_class(calls)}
    )

    export = next(call for call in calls if call[0] == "ExportAsFixedFormat")
    assert export[1][1] == 17
    assert Path(export[1][0]).name == "converted.pdf"
    assert result == request.output
    assert calls[0] == ("open_document", str(request.source), True, False)
    assert calls[-1] == ("close", False)


def test_spreadsheet_exports_the_whole_workbook(tmp_path: Path):
    calls = []
    request = _request(tmp_path, "xlsx", "spreadsheet")

    result = convert_windows(
        request, composer_classes={"spreadsheet": _composer_class(calls)}
    )

    export = next(call for call in calls if call[0] == "ExportAsFixedFormat")
    assert export[1][0] == 0
    assert Path(export[1][1]).name == "converted.pdf"
    assert result == request.output
    assert calls[-1] == ("close", False)


def test_presentation_uses_save_as_pdf_format(tmp_path: Path):
    calls = []
    request = _request(tmp_path, "pptx", "presentation")

    result = convert_windows(
        request, composer_classes={"presentation": _composer_class(calls)}
    )

    save = next(call for call in calls if call[0] == "SaveAs")
    assert Path(save[1][0]).name == "converted.pdf"
    assert save[1][1] == 32
    assert result == request.output
    assert calls[-1] == ("close", False)


def test_windows_export_failure_closes_source_and_publishes_nothing(
    tmp_path: Path,
):
    calls = []
    request = _request(tmp_path, "doc", "writer")

    class FailingDocument(FakeDocument):
        def ExportAsFixedFormat(self, *args):
            self.calls.append(("ExportAsFixedFormat", args))
            raise RuntimeError("COM export failed")

    class FailingComposer(FakeComposer):
        def __init__(self, calls, source, read_only, visible):
            super().__init__(calls, source, read_only, visible)
            self.doc = FailingDocument(calls)

    class FailingClass:
        @classmethod
        def open_document(cls, source, read_only=False, visible=False):
            return FailingComposer(calls, source, read_only, visible)

    with pytest.raises(RuntimeError, match="COM export failed"):
        convert_windows(request, composer_classes={"writer": FailingClass})

    assert calls[-1] == ("close", False)
    assert not request.output.exists()


def test_windows_backend_rejects_unknown_component(tmp_path: Path):
    request = _request(tmp_path, "docx", "unknown")
    with pytest.raises(ValueError, match="Unsupported conversion component"):
        convert_windows(request, composer_classes={})
