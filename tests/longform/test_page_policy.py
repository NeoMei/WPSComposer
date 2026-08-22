from __future__ import annotations

import pytest

from skills.WPSComposer.scripts.document_model import (
    AbstractBlock,
    KeywordsBlock,
    Paragraph,
    Section,
    Span,
    StructuredDocument,
)
from skills.WPSComposer.scripts.longform.policy import (
    LongformPolicy,
    build_policy,
)
from skills.WPSComposer.scripts.longform.semantic import LongformConfig
from skills.WPSComposer.scripts.longform.unicode_text import display_units
from skills.WPSComposer.scripts.longform.page_policy import (
    PageSkeletonPolicy,
    build_page_policy,
)


def _policy_from_config(config: LongformConfig) -> LongformPolicy:
    return build_policy(config)


def _role_names(policy: PageSkeletonPolicy) -> list[str]:
    return [section.role for section in policy.sections]


def _make_landscape_section(heading: str) -> Section:
    section = Section(level=1, heading=heading)
    section.orientation = "landscape"  # type: ignore[attr-defined]
    return section


# ---------------------------------------------------------------------------
# Title page
# ---------------------------------------------------------------------------

def test_title_page_omitted_when_title_is_empty() -> None:
    config = LongformConfig(title="", author="Author", title_page=True)
    doc = StructuredDocument(title="", sections=[Section(level=1, heading="Intro")])
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    assert "cover" not in _role_names(skeleton)


def test_title_page_present_when_author_or_date_exists() -> None:
    config = LongformConfig(title="My Report", author="Author", date="", title_page=True)
    doc = StructuredDocument(title="My Report", sections=[Section(level=1, heading="Intro")])
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    assert skeleton.sections[0].role == "cover"
    assert skeleton.sections[0].has_header is False
    assert skeleton.sections[0].has_footer is False
    assert skeleton.sections[0].page_number_format == "none"
    assert skeleton.sections[0].start_page_number is None


# ---------------------------------------------------------------------------
# Front matter
# ---------------------------------------------------------------------------

def test_front_matter_section_for_abstract_keywords_and_toc() -> None:
    config = LongformConfig(title="Report", author="A", toc=True, figure_index=False, table_index=False)
    doc = StructuredDocument(
        title="Report",
        abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
        keywords=KeywordsBlock(keywords=["one", "two"]),
        sections=[Section(level=1, heading="Intro")],
    )
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    assert "front_matter" in _role_names(skeleton)
    front = next(s for s in skeleton.sections if s.role == "front_matter")
    assert front.page_number_format == "roman"
    assert front.start_page_number == 1
    assert front.restart_numbering is True
    assert front.has_header is False
    assert front.has_footer is True
    assert front.includes_abstract is True
    assert front.includes_keywords is True
    assert front.includes_toc is True


# ---------------------------------------------------------------------------
# Body numbering
# ---------------------------------------------------------------------------

def test_body_section_restarts_arabic_numbering() -> None:
    config = LongformConfig(title="Report", author="A", toc=True)
    doc = StructuredDocument(
        title="Report",
        abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
        sections=[Section(level=1, heading="Intro")],
    )
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    body = next(s for s in skeleton.sections if s.role == "body")
    assert body.page_number_format == "arabic"
    assert body.start_page_number == 1
    assert body.restart_numbering is True
    assert body.has_header is True
    assert body.has_footer is True
    assert body.link_to_previous_header is False
    assert body.link_to_previous_footer is False


def test_no_front_matter_body_starts_arabic_immediately() -> None:
    config = LongformConfig(title="Report", author="A", toc=False)
    doc = StructuredDocument(
        title="Report",
        sections=[Section(level=1, heading="Intro"), Section(level=1, heading="Next")],
    )
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    assert _role_names(skeleton) == ["body"]
    body = skeleton.sections[0]
    assert body.page_number_format == "arabic"
    assert body.start_page_number == 1
    assert body.restart_numbering is True
    assert body.link_to_previous_header is False
    assert body.link_to_previous_footer is False


# ---------------------------------------------------------------------------
# Temporary landscape ordering
# ---------------------------------------------------------------------------

def test_temporary_landscape_continues_numbering() -> None:
    config = LongformConfig(title="Report", author="A")
    body_section = Section(level=1, heading="Body")
    landscape_section = _make_landscape_section("Wide Table")
    doc = StructuredDocument(
        title="Report",
        sections=[body_section, landscape_section],
    )
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    landscape = next(s for s in skeleton.sections if s.role == "landscape")
    assert landscape.page_number_format == "continue"
    assert landscape.restart_numbering is False
    assert landscape.link_to_previous_header is True
    assert landscape.link_to_previous_footer is True


def test_multiple_landscape_sections_preserve_document_order() -> None:
    config = LongformConfig(title="Report", author="A")
    doc = StructuredDocument(
        title="Report",
        sections=[
            Section(level=1, heading="P1"),
            _make_landscape_section("L1"),
            Section(level=1, heading="P2"),
            _make_landscape_section("L2"),
        ],
    )
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    roles = _role_names(skeleton)
    assert roles == ["body", "landscape", "body", "landscape"]

    # First body starts fresh; subsequent content sections link to the previous one.
    assert skeleton.sections[0].link_to_previous_header is False
    assert skeleton.sections[0].link_to_previous_footer is False
    for section in skeleton.sections[1:]:
        assert section.link_to_previous_header is True
        assert section.link_to_previous_footer is True


def test_document_starting_with_landscape_is_not_reordered() -> None:
    config = LongformConfig(title="Report", author="A")
    doc = StructuredDocument(
        title="Report",
        sections=[
            _make_landscape_section("First Wide"),
            Section(level=1, heading="Then Portrait"),
        ],
    )
    policy = _policy_from_config(config)
    skeleton = build_page_policy(doc, config, policy)

    assert _role_names(skeleton) == ["landscape", "body"]
    assert skeleton.sections[0].link_to_previous_header is False
    assert skeleton.sections[0].link_to_previous_footer is False
    assert skeleton.sections[1].link_to_previous_header is True
    assert skeleton.sections[1].link_to_previous_footer is True


# ---------------------------------------------------------------------------
# Header shortening
# ---------------------------------------------------------------------------

def test_header_shortened_to_64_display_units_with_overflow_fallback_at_32() -> None:
    long_header = "国" * 50  # 100 display units
    config = LongformConfig(title="Report", author="A", header=long_header)
    doc = StructuredDocument(title="Report", sections=[Section(level=1, heading="Intro")])
    policy = _policy_from_config(config)

    assert display_units(policy.header_text) <= 64
    assert display_units(policy.header_overflow_text) <= 32

    skeleton = build_page_policy(doc, config, policy)
    body = next(s for s in skeleton.sections if s.role == "body")
    assert display_units(body.header_text) <= 64


def test_header_exactly_64_and_32_display_units_with_emoji_grapheme_clusters() -> None:
    # Each emoji is 2 display units. 31 emoji = 62 units, plus ellipsis = 64.
    long_header = "😀" * 40  # 80 display units
    config = LongformConfig(title="Report", author="A", header=long_header)
    policy = _policy_from_config(config)

    assert display_units(policy.header_text) == 64
    assert display_units(policy.header_overflow_text) == 32


# ---------------------------------------------------------------------------
# TOC density minima
# ---------------------------------------------------------------------------

def test_toc_density_minima_present() -> None:
    config = LongformConfig(title="Report", author="A", toc=True)
    policy = _policy_from_config(config)

    density = policy.toc_density
    assert density["min_font_size_pt"]["toc1"] == 10.5
    assert density["min_font_size_pt"]["toc2"] == 10.0
    assert density["min_font_size_pt"]["toc3"] == 10.0
    assert density["min_space_before_pt"]["toc1"] == 0.0
    assert density["min_space_after_pt"]["toc3"] == 0.0


# ---------------------------------------------------------------------------
# Determinism and degradation
# ---------------------------------------------------------------------------

def test_build_page_policy_is_deterministic() -> None:
    config = LongformConfig(title="Report", author="A", toc=True)
    doc = StructuredDocument(
        title="Report",
        abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
        sections=[
            Section(level=1, heading="Intro"),
            _make_landscape_section("Wide"),
            Section(level=1, heading="Outro"),
        ],
    )
    policy = _policy_from_config(config)

    first = build_page_policy(doc, config, policy)
    second = build_page_policy(doc, config, policy)
    assert first == second


def test_build_page_policy_never_raises_on_bad_input() -> None:
    # Malformed/None inputs must degrade to an empty skeleton rather than propagate.
    result = build_page_policy(None, None, None)  # type: ignore[arg-type]
    assert isinstance(result, PageSkeletonPolicy)
    assert result.sections == ()
