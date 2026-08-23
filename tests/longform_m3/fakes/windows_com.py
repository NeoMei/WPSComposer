from __future__ import annotations

from pathlib import Path
from typing import Any

from skills.WPSComposer.scripts.longform.executor import FieldSnapshot


class NativeFailure(RuntimeError):
    def __init__(self, code: str) -> None:
        super().__init__("native failure")
        self.code = code


class RecordingNativeComposer:
    def __init__(self) -> None:
        self.calls: list[tuple[str, dict[str, Any]]] = []
        self.closed = False
        self.failures: dict[str, str] = {}
        self.resource_paths_seen: list[str] = []

    def _call(self, name: str, **kwargs: Any) -> dict[str, Any] | None:
        code = self.failures.get(name)
        if code:
            raise NativeFailure(code)
        self.calls.append((name, kwargs))
        for path in (kwargs.get("resource_locators") or {}).values():
            assert Path(path).is_file()
            self.resource_paths_seen.append(path)
        return None

    def close(self, save_changes: bool = False) -> None:
        self.closed = True

    def save_docx(self, path: str) -> str:
        self._call("save_docx", path=path)
        return path

    def add_captioned_figure_native(self, **kwargs: Any) -> dict[str, Any] | None:
        return self._call("figure", **kwargs)

    def add_semantic_table_native(self, **kwargs: Any) -> dict[str, Any] | None:
        return self._call("table", **kwargs)

    def add_equation_number_native(self, **kwargs: Any) -> dict[str, Any] | None:
        return self._call("equation", **kwargs)

    def add_cross_reference_paragraph(self, **kwargs: Any) -> dict[str, Any] | None:
        return self._call("reference", **kwargs)

    def insert_caption_index_native(self, **kwargs: Any) -> dict[str, Any] | None:
        return self._call("index", **kwargs)

    def repaginate_and_update_numbering(self) -> None:
        self._call("phase-numbering")

    def refresh_bookmarks_and_references(self) -> None:
        self._call("phase-references")

    def refresh_indexes(self) -> None:
        self._call("phase-indexes")

    def repaginate_and_update_page_fields(self) -> None:
        self._call("phase-pages")

    def snapshot_fields(self):
        self._call("phase-snapshot")
        return (
            FieldSnapshot(
                stable_key=("doc:finalize", "PAGE", 0),
                field_category="page",
                result_hash="0" * 64,
                toc_page_count=0,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=1,
            ),
        )

    def upsert_document_quality_notice(self, issue) -> None:
        self._call("quality-notice", code=issue.code)

    def add_degradation_notice(self, **kwargs: Any) -> None:
        self._call("notice", **kwargs)

    def add_inline_degradation(self, **kwargs: Any) -> None:
        self._call("inline", **kwargs)

    def configure_section(self, **kwargs: Any) -> None:
        self._call("section", **kwargs)
