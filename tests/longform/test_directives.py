from __future__ import annotations

import unicodedata

import pytest

from skills.WPSComposer.scripts.longform.directives import (
    DIRECTIVE_SYNTAX_INVALID,
    DIRECTIVE_UNCLOSED,
    NESTED_DIRECTIVE_UNSUPPORTED,
    BlockDirective,
    scan_block_directives,
)


def _only_directive(markdown: str) -> BlockDirective:
    regions = scan_block_directives(markdown)

    assert len(regions) == 1
    directive = regions[0]
    assert isinstance(directive, BlockDirective)
    return directive


def test_scans_identifier_json_strings_unquoted_tokens_and_markdown_regions() -> None:
    markdown = (
        "Before\n\n"
        ':::figure {#fig:model caption="A \\\"quoted\\\" \\/ path" '
        "width=360pt source=assets/a%20b.png}\n"
        "![model](assets/model.png)\n"
        ":::\n"
        "\nAfter\n"
    )

    regions = scan_block_directives(markdown)

    assert regions[0] == "Before\n\n"
    directive = regions[1]
    assert isinstance(directive, BlockDirective)
    assert directive.name == "figure"
    assert directive.identifier == "fig:model"
    assert directive.attributes == {
        "caption": 'A "quoted" / path',
        "width": "360pt",
        "source": "assets/a%20b.png",
    }
    assert directive.body == "![model](assets/model.png)\n"
    assert directive.start_line == 3
    assert directive.issues == ()
    assert regions[2] == "\nAfter\n"


def test_normalizes_decoded_attribute_values_to_nfc() -> None:
    decomposed = "Cafe\u0301"

    directive = _only_directive(
        f':::figure {{caption="{decomposed}" label="{decomposed}"}}\n:::\n'
    )

    assert directive.attributes == {
        "caption": unicodedata.normalize("NFC", decomposed),
        "label": unicodedata.normalize("NFC", decomposed),
    }


@pytest.mark.parametrize(
    "attributes",
    [
        "#fig:first #fig:second",
        "caption=first caption=second",
    ],
    ids=("duplicate-id", "duplicate-key"),
)
def test_duplicate_identifiers_or_keys_discard_all_partial_attributes(
    attributes: str,
) -> None:
    directive = _only_directive(
        f":::figure {{{attributes}}}\nReadable body\n:::\n"
    )

    assert directive.identifier is None
    assert directive.attributes == {}
    assert directive.body == "Readable body\n"
    assert directive.issues == (DIRECTIVE_SYNTAX_INVALID,)


@pytest.mark.parametrize(
    "attribute",
    [
        r'caption="bad\q"',
        r'caption="line\nfeed"',
        r'caption="nul\u0000value"',
        r'caption="surrogate\uD800value"',
        'caption="raw\x85control"',
        'caption="unterminated',
    ],
    ids=(
        "unknown-escape",
        "escaped-line-feed",
        "escaped-nul",
        "lone-surrogate",
        "raw-c1-control",
        "unterminated-string",
    ),
)
def test_invalid_string_attributes_degrade_without_raising(attribute: str) -> None:
    directive = _only_directive(
        f":::figure {{#fig:model width=full {attribute}}}\nBody stays readable\n:::\n"
    )

    assert directive.identifier is None
    assert directive.attributes == {}
    assert directive.body == "Body stays readable\n"
    assert directive.issues == (DIRECTIVE_SYNTAX_INVALID,)


@pytest.mark.parametrize(
    "opening",
    [
        ":::Figure",
        ":::figure trailing",
        ":::figure {Bad=value}",
        ":::figure {caption=Chinese中文}",
        ":::figure {caption=}",
    ],
    ids=(
        "uppercase-name",
        "text-after-name",
        "uppercase-key",
        "non-ascii-token",
        "empty-token",
    ),
)
def test_invalid_opening_syntax_preserves_readable_body(opening: str) -> None:
    directive = _only_directive(f"{opening}\nReadable body\n:::\n")

    assert directive.attributes == {}
    assert directive.body == "Readable body\n"
    assert directive.issues == (DIRECTIVE_SYNTAX_INVALID,)


def test_nested_directive_is_preserved_as_text_in_the_outer_body() -> None:
    directive = _only_directive(
        ":::abstract\n"
        "Before nested\n"
        ":::keywords\n"
        "one; two\n"
        ":::\n"
        "After nested\n"
        ":::\n"
    )

    assert directive.name == "abstract"
    assert directive.body == (
        "Before nested\n"
        ":::keywords\n"
        "one; two\n"
        ":::\n"
        "After nested\n"
    )
    assert directive.issues == (NESTED_DIRECTIVE_UNSUPPORTED,)


def test_unclosed_directive_keeps_all_following_content_in_its_readable_body() -> None:
    directive = _only_directive(
        ":::abstract\nFirst paragraph\n\n# Later heading\nLast paragraph\n"
    )

    assert directive.body == "First paragraph\n\n# Later heading\nLast paragraph\n"
    assert directive.start_line == 1
    assert directive.issues == (DIRECTIVE_UNCLOSED,)


@pytest.mark.parametrize("fence", ["```", "~~~~"])
def test_directive_markers_inside_fenced_code_are_plain_markdown(fence: str) -> None:
    fenced = (
        f"{fence}markdown\n"
        ":::figure {#fig:code caption=literal}\n"
        "inside code\n"
        ":::\n"
        f"{fence}\n\n"
    )
    markdown = fenced + ":::page-break\n:::\n"

    regions = scan_block_directives(markdown)

    assert regions[0] == fenced
    assert isinstance(regions[1], BlockDirective)
    assert regions[1].name == "page-break"


def test_indented_code_block_quotes_and_list_contents_are_plain_markdown() -> None:
    literal = (
        "    :::figure\n"
        "    indented code\n"
        "    :::\n\n"
        "> :::figure\n"
        "> quoted\n"
        "> :::\n\n"
        "- :::figure\n"
        "  listed\n"
        "  :::\n\n"
    )
    markdown = literal + ":::keywords\none; two\n:::\n"

    regions = scan_block_directives(markdown)

    assert regions[0] == literal
    assert isinstance(regions[1], BlockDirective)
    assert regions[1].name == "keywords"


def test_fenced_code_inside_a_directive_does_not_look_nested() -> None:
    directive = _only_directive(
        ":::abstract\n"
        "```markdown\n"
        ":::figure\n"
        ":::\n"
        "```\n"
        ":::\n"
    )

    assert directive.body == "```markdown\n:::figure\n:::\n```\n"
    assert directive.issues == ()


def test_unknown_but_lexically_valid_name_is_not_a_syntax_error() -> None:
    directive = _only_directive(":::future-block {mode=preview}\nBody\n:::\n")

    assert directive.name == "future-block"
    assert directive.attributes == {"mode": "preview"}
    assert directive.issues == ()
