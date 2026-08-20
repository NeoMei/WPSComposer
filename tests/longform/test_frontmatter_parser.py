from __future__ import annotations

import json
import unicodedata

import pytest

from skills.WPSComposer.scripts.longform.frontmatter_parser import (
    FRONTMATTER_INVALID,
    FRONTMATTER_UNCLOSED,
    parse_frontmatter_document,
)


def _document(frontmatter: str, body: str = "# Body\n") -> str:
    return f"---\n{frontmatter}---\n{body}"


def test_parses_safe_yaml_data_subset_and_preserves_body() -> None:
    text = _document(
        """title: 决策即服务
date: 2026-08-20
enabled: true
missing: null
count: 12
ratio: 1.25
authors:
  - 张三
  - 李四
options:
  toc: false
  levels: [1, 2, 3]
labels: {primary: 蓝色, secondary: "绿色"}
""",
        "# 正文\n",
    )

    result = parse_frontmatter_document(text)

    assert result.values == {
        "title": "决策即服务",
        "date": "2026-08-20",
        "enabled": True,
        "missing": None,
        "count": 12,
        "ratio": 1.25,
        "authors": ["张三", "李四"],
        "options": {"toc": False, "levels": [1, 2, 3]},
        "labels": {"primary": "蓝色", "secondary": "绿色"},
    }
    assert result.issues == ()
    assert result.boundary == (0, text.index("# 正文"))
    assert result.body == "# 正文\n"
    json.dumps(result.values, ensure_ascii=False, allow_nan=False)


def test_document_without_frontmatter_is_unchanged() -> None:
    text = "# Body\n\n---\nnot frontmatter\n"

    result = parse_frontmatter_document(text)

    assert result.values == {}
    assert result.issues == ()
    assert result.boundary is None
    assert result.body == text


def test_accepts_a_flow_mapping_at_the_document_root() -> None:
    result = parse_frontmatter_document(
        _document("{title: Draft, options: {toc: true}}\n")
    )

    assert result.values == {"title": "Draft", "options": {"toc": True}}
    assert result.issues == ()


def test_single_quoted_hash_and_doubled_quote_are_not_comments() -> None:
    result = parse_frontmatter_document(
        _document("title: 'It''s # safe' # actual comment\n")
    )

    assert result.values == {"title": "It's # safe"}
    assert result.issues == ()


def test_unclosed_frontmatter_preserves_entire_document() -> None:
    text = "---\ntitle: Draft\n# Still source text\n"

    result = parse_frontmatter_document(text)

    assert result.values == {}
    assert result.issues == (FRONTMATTER_UNCLOSED,)
    assert result.boundary is None
    assert result.body == text


def test_closing_boundary_must_be_within_64_kib_utf8() -> None:
    text = "---\nvalue: " + ("界" * 22_000) + "\n---\nBody\n"

    result = parse_frontmatter_document(text)

    assert result.values == {}
    assert result.issues == (FRONTMATTER_UNCLOSED,)
    assert result.boundary is None
    assert result.body == text


def test_closing_marker_may_end_exactly_at_64_kib() -> None:
    prefix = "title: value\n#"
    padding = (64 * 1024) - len(prefix.encode("utf-8")) - len("\n---")
    text = _document(prefix + ("x" * padding) + "\n", "Body\n")

    result = parse_frontmatter_document(text)

    assert result.values == {"title": "value"}
    assert result.issues == ()
    assert result.body == "Body\n"


@pytest.mark.parametrize(
    "frontmatter",
    [
        "title: first\ntitle: second\n",
        "title: &shared value\ncopy: *shared\n",
        "base: &base {toc: true}\nconfig:\n  <<: *base\n",
        "title: !unknown value\n",
        "title: !!python/object/apply:os.system [echo, unsafe]\n",
        "? [compound, key]\n: value\n",
        "1: numeric key\n",
    ],
    ids=(
        "duplicate-key",
        "anchor-alias",
        "merge-key",
        "unknown-tag",
        "python-object-tag",
        "compound-key",
        "non-string-key",
    ),
)
def test_rejects_unsafe_or_non_string_mapping_constructs(
    frontmatter: str,
) -> None:
    result = parse_frontmatter_document(_document(frontmatter))

    assert result.values == {}
    assert result.issues == (FRONTMATTER_INVALID,)
    assert result.body == "# Body\n"


@pytest.mark.parametrize(
    "frontmatter",
    [
        'title: "unsafe\\ntext"\n',
        'title: "unsafe\\u0000text"\n',
        'title: "unsafe\x85text"\n',
        'title: "unsafe\ud800text"\n',
        '"unsafe\\u0000key": value\n',
    ],
    ids=(
        "escaped-line-feed",
        "escaped-nul",
        "c1-control",
        "lone-surrogate",
        "control-in-key",
    ),
)
def test_rejects_controls_and_non_scalar_unicode(frontmatter: str) -> None:
    result = parse_frontmatter_document(_document(frontmatter))

    assert result.values == {}
    assert result.issues == (FRONTMATTER_INVALID,)


def test_enforces_maximum_depth_of_eight() -> None:
    def nested_mapping(depth: int) -> str:
        lines = [
            ("  " * level) + f"level{level + 1}:"
            for level in range(depth - 1)
        ]
        lines.append(("  " * (depth - 1)) + f"level{depth}: value")
        return "\n".join(lines) + "\n"

    valid = parse_frontmatter_document(_document(nested_mapping(8)))
    invalid = parse_frontmatter_document(_document(nested_mapping(9)))

    assert valid.issues == ()
    assert invalid.values == {}
    assert invalid.issues == (FRONTMATTER_INVALID,)


def test_enforces_256_total_mapping_entries_or_list_elements() -> None:
    valid_mapping = "".join(f"key_{index}: value\n" for index in range(256))
    invalid_mapping = valid_mapping + "overflow: value\n"
    valid_list = "items:\n" + "".join("  - value\n" for _ in range(255))
    invalid_list = valid_list + "  - overflow\n"

    assert parse_frontmatter_document(_document(valid_mapping)).issues == ()
    assert parse_frontmatter_document(_document(valid_list)).issues == ()
    assert parse_frontmatter_document(_document(invalid_mapping)).issues == (
        FRONTMATTER_INVALID,
    )
    assert parse_frontmatter_document(_document(invalid_list)).issues == (
        FRONTMATTER_INVALID,
    )


def test_normalizes_all_string_keys_and_values_to_nfc() -> None:
    decomposed_key = "résumé"
    decomposed_value = "Café 作品"
    text = _document(
        f'"{decomposed_key}":\n'
        f'  - "{decomposed_value}"\n'
        f'  - nested: "{decomposed_value}"\n'
    )

    result = parse_frontmatter_document(text)

    normalized_key = unicodedata.normalize("NFC", decomposed_key)
    normalized_value = unicodedata.normalize("NFC", decomposed_value)
    assert result.values == {
        normalized_key: [normalized_value, {"nested": normalized_value}]
    }
    assert result.issues == ()


def test_duplicate_keys_are_detected_after_nfc_normalization() -> None:
    result = parse_frontmatter_document(
        _document('"Café": first\n"Café": second\n')
    )

    assert result.values == {}
    assert result.issues == (FRONTMATTER_INVALID,)
