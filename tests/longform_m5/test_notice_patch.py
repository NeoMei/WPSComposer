from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
from skills.WPSComposer.scripts.longform.quality import (
    QualityConfidence,
    QualityFinding,
    QualitySeverity,
)
from skills.WPSComposer.scripts.macos_probe.models import (
    ProbeResult,
    ProtocolError,
    validate_longform_notice_patch_request,
    validate_longform_notice_patch_value,
)


BOOKMARK = "wpsc_fig_" + "a" * 24
HEADING_BOOKMARK = "wpsc_head_" + "b" * 24


def _notice():
    return QualityFinding(
        code="IMAGE_LOW_DPI",
        severity=QualitySeverity.DEGRADED,
        confidence=QualityConfidence.HIGH,
        message="Image remains below the configured DPI threshold",
        node_id="fig:one",
        page=2,
        repair_key="low-dpi-notice",
    )


def _request():
    return {
        "sourcePath": "/private/source.docx",
        "outputPath": "/private/output.docx",
        "bookmarks": [
            {"nodeId": "fig:one", "bookmarkName": BOOKMARK},
            {"nodeId": "head:one", "bookmarkName": HEADING_BOOKMARK},
        ],
        "notices": [
            {
                "code": "IMAGE_LOW_DPI",
                "message": "Image remains below the configured DPI threshold",
                "fallback": "low-dpi-notice",
                "placement": "block",
                "nodeId": "fig:one",
                "page": 2,
                "bookmarkName": BOOKMARK,
            }
        ],
    }


def _value(output="/private/output.docx"):
    return {
        "outputPath": output,
        "appliedNotices": 1,
        "issueCodes": [],
        "fieldSnapshots": [],
        "paginationMap": {
            "version": "M5-v1",
            "nodes": [
                {
                    "nodeId": "fig:one",
                    "story": "main",
                    "sections": ["body"],
                    "pageStart": 2,
                    "pageEnd": 2,
                    "range": "10:20",
                    "fragments": [{"page": 2, "bounds": [72, 90, 300, 240]}],
                }
            ],
        },
    }


def test_notice_patch_request_and_value_are_closed_roundtrips():
    assert validate_longform_notice_patch_request(_request()) == _request()
    assert validate_longform_notice_patch_value(_value()) == _value()


@pytest.mark.parametrize(
    "change",
    [
        {"extra": True},
        {"notices": []},
        {"notices": [{**_request()["notices"][0], "page": 0}]},
        {"notices": [{**_request()["notices"][0], "bookmarkName": "unsafe"}]},
        {"notices": [{**_request()["notices"][0], "message": "/Users/private/a"}]},
        {"notices": [{**_request()["notices"][0], "fallback": "../../escape"}]},
    ],
)
def test_notice_patch_request_rejects_unknown_private_or_unbounded_data(change):
    request = {**_request(), **change}
    with pytest.raises(ProtocolError, match="notice patch request is invalid"):
        validate_longform_notice_patch_request(request)


@pytest.mark.parametrize(
    "value",
    [
        {**_value(), "private": "/Users/private/a"},
        {**_value(), "appliedNotices": -1},
        {**_value(), "paginationMap": {"version": "M5-v1", "nodes": [{"nodeId": "/private/a", "fragments": []}]}},
        {**_value(), "paginationMap": {"version": "M5-v1", "nodes": [{"nodeId": "fig:one", "fragments": [{"page": 1, "bounds": [0, 0, 0, 1]}]}]}},
    ],
)
def test_notice_patch_value_rejects_invalid_or_private_response(value):
    with pytest.raises(ProtocolError, match="notice patch result is invalid"):
        validate_longform_notice_patch_value(value)


class _Bridge:
    def __init__(self):
        self.params = None

    def issue(self, component, method, params):
        assert (component, method) == ("writer", "patch_longform_quality_notices")
        self.params = params
        return SimpleNamespace(id="patch-1")

    def wait_result(self, command_id, timeout):
        assert command_id == "patch-1"
        return ProbeResult(command_id, True, _value(self.params["outputPath"]), None)


def test_macos_executor_patches_once_and_returns_fresh_m5_pagination(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"PK\x03\x04" + b"x" * 300)
    bridge = _Bridge()
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path))

    outcome = executor.patch_quality_notices(
        source,
        (_notice(),),
        {"fig:one": BOOKMARK},
        deadline=None,
    )

    assert bridge.params["notices"][0]["bookmarkName"] == BOOKMARK
    assert bridge.params["bookmarks"] == [
        {"nodeId": "fig:one", "bookmarkName": BOOKMARK}
    ]
    assert bridge.params["notices"][0]["page"] == 2
    assert outcome.pagination_map.version == "M5-v1"
    assert outcome.pagination_map.nodes[0].fragments[0].bounds == (72, 90, 300, 240)


def test_macos_executor_requires_mapped_bookmark_for_block_notice(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"PK\x03\x04" + b"x" * 300)
    executor = MacOSLongformExecutor(bridge=_Bridge(), staging_dir=str(tmp_path))
    with pytest.raises(ValueError, match="bookmark"):
        executor.patch_quality_notices(source, (_notice(),), {}, deadline=None)


class _WindowsPatchComposer:
    def __init__(self):
        self.opened = None
        self.notices = []
        self.saved = None
        self.closed = False
        self.refresh_rounds = 0

    def open_existing_for_patch(self, source):
        self.opened = source

    def add_quality_notice_at_bookmark(self, **notice):
        self.notices.append(notice)

    def pagination_fragment_for_bookmark(self, node_id, bookmark):
        return {
            "nodeId": node_id,
            "story": "main",
            "sections": ["body"],
            "pageStart": 2,
            "pageEnd": 2,
            "range": "10:20",
            "fragments": [{"page": 2, "bounds": [72, 90, 300, 240]}],
        }

    def refresh_fields(self, round_index):
        from skills.WPSComposer.scripts.longform.executor import FieldSnapshot

        self.refresh_rounds += 1
        return (
            FieldSnapshot(
                stable_key=("doc:finalize", "PAGE", 0),
                field_category="page",
                result_hash="stable",
                toc_page_count=0,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=2,
            ),
        )

    def upsert_document_quality_notice(self, issue):
        raise AssertionError("stable patch must not insert instability notice")

    def save_docx(self, target):
        self.saved = target
        Path(target).write_bytes(b"PK\x03\x04" + b"x" * 300)

    def close(self, save_changes=False):
        self.closed = True


def test_windows_executor_patches_in_dedicated_composer_and_refreshes(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"PK\x03\x04" + b"x" * 300)
    composer = _WindowsPatchComposer()
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path), composer_factory=lambda: composer
    )
    outcome = executor.patch_quality_notices(
        source,
        (_notice(),),
        {"fig:one": BOOKMARK, "head:one": HEADING_BOOKMARK},
        deadline=None,
    )
    assert composer.opened == str(source.resolve())
    assert composer.notices[0]["bookmark_name"] == BOOKMARK
    assert composer.refresh_rounds == 2
    assert composer.closed is True
    assert outcome.pagination_map.version == "M5-v1"
    assert [node.node_id for node in outcome.pagination_map.nodes] == [
        "fig:one", "head:one"
    ]
    assert outcome.applied_operations == 1
