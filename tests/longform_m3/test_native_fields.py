from __future__ import annotations

import pytest

from skills.WPSComposer.scripts.document_model import CaptionBinding, CrossReferenceRun


def test_caption_descriptors_are_closed_and_object_local() -> None:
    from skills.WPSComposer.scripts.longform.native_fields import (
        caption_numbering_descriptor,
    )

    global_binding = CaptionBinding("global", None, "wpsc_fig_" + "a" * 24, True, True)
    chapter_binding = CaptionBinding("chapter", "sec:one", "wpsc_tab_" + "b" * 24, True, True)

    assert caption_numbering_descriptor("figure", global_binding) == {
        "mode": "global",
        "sequenceId": "WPSC_FIG",
        "chapterStyleLevel": None,
        "resetLevel": None,
        "prefix": "图 ",
        "suffix": "",
    }
    assert caption_numbering_descriptor("table", chapter_binding) == {
        "mode": "chapter",
        "sequenceId": "WPSC_TAB",
        "chapterStyleLevel": 1,
        "resetLevel": 1,
        "prefix": "表 ",
        "suffix": "",
    }


def test_equation_descriptor_has_number_shell_only() -> None:
    from skills.WPSComposer.scripts.longform.native_fields import (
        caption_numbering_descriptor,
    )

    descriptor = caption_numbering_descriptor(
        "equation", CaptionBinding("chapter", "sec:one", "wpsc_eq_" + "c" * 24, True, True)
    )
    assert descriptor == {
        "mode": "chapter",
        "sequenceId": "WPSC_EQ",
        "chapterStyleLevel": 1,
        "resetLevel": 1,
        "prefix": "(",
        "suffix": ")",
    }
    assert "latex" not in descriptor
    assert "fieldCode" not in descriptor


def test_cross_reference_descriptor_excludes_external_id_and_user_text() -> None:
    from skills.WPSComposer.scripts.longform.native_fields import (
        cross_reference_descriptor,
    )

    run = CrossReferenceRun(
        node_id="ref:1",
        target_id="fig:user SEQ EVIL",
        target_node_id="fig:target",
        target_kind="figure",
        bookmark_name="wpsc_fig_" + "d" * 24,
        fallback_text="user REF injection",
    )
    assert cross_reference_descriptor(run) == {
        "targetNodeId": "fig:target",
        "targetKind": "figure",
        "bookmarkName": "wpsc_fig_" + "d" * 24,
        "prefix": "图 ",
        "suffix": "",
    }


@pytest.mark.parametrize("kind", ["caption", "custom", "WPSC_FIG"])
def test_native_descriptors_reject_uncontrolled_kinds(kind: str) -> None:
    from skills.WPSComposer.scripts.longform.native_fields import (
        caption_numbering_descriptor,
    )

    with pytest.raises(ValueError):
        caption_numbering_descriptor(
            kind, CaptionBinding("global", None, None, False, False)
        )


def test_cross_reference_descriptor_requires_resolved_target_and_bookmark() -> None:
    from skills.WPSComposer.scripts.longform.native_fields import (
        cross_reference_descriptor,
    )

    unresolved = CrossReferenceRun("ref:1", "fig:x", None, "figure", None, "[图]")
    with pytest.raises(ValueError):
        cross_reference_descriptor(unresolved)
