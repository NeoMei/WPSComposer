from __future__ import annotations

import unicodedata

import pytest

from skills.WPSComposer.scripts.longform.unicode_text import (
    contains_han,
    display_units,
    normalize_visible_text,
    shorten_display_units,
)


def test_normalize_visible_text_combines_decomposed_diacritics() -> None:
    decomposed = "Cafe\u0301"  # Cafe + combining acute
    nfc = unicodedata.normalize("NFC", decomposed)

    assert normalize_visible_text(decomposed) == nfc
    assert len(normalize_visible_text(decomposed)) == 4


def test_normalize_visible_text_is_stable_for_already_nfc() -> None:
    text = "已经NFC的文本"

    assert normalize_visible_text(text) == text


def test_contains_han_detects_basic_and_extension_characters() -> None:
    assert contains_han("中文") is True
    assert contains_han("CJK Extension G: \U00030000") is True
    assert contains_han("\U0003134F末") is True
    assert contains_han("English 123") is False
    assert contains_han("") is False


def test_contains_han_distinguishes_script_han_from_cjk_kana_and_hangul() -> None:
    assert contains_han("中") is True
    assert contains_han("永") is True
    assert contains_han("あ") is False  # Hiragana has its own script
    assert contains_han("カ") is False  # Katakana has its own script
    assert contains_han("한") is False  # Hangul has its own script


def test_display_units_counts_ascii_as_one_per_cluster() -> None:
    assert display_units("Hello") == 5


def test_display_units_counts_cjk_as_two() -> None:
    assert display_units("中文") == 4
    assert display_units("\U00030000") == 2  # Extension G


def test_display_units_counts_combining_marks_without_splitting() -> None:
    # "e" + combining acute is one cluster, one unit in Latin context
    decomposed = "Cafe\u0301"
    assert display_units(decomposed) == 5
    assert display_units(unicodedata.normalize("NFC", decomposed)) == 5


def test_display_units_keeps_zwj_emoji_sequences_together() -> None:
    # Family ZWJ sequence: one grapheme cluster
    family = "\U0001F468\u200D\U0001F469\u200D\U0001F467\u200D\U0001F466"
    assert display_units(family) == 2

    # Man astronaut: one cluster
    astronaut = "\U0001F468\u200D\U0001F680"
    assert display_units(astronaut) == 2


def test_display_units_pairs_regional_indicators() -> None:
    flag = "\U0001F1E8\U0001F1F3"  # CN flag
    assert display_units(flag) == 2

    two_flags = flag + "\U0001F1FA\U0001F1F8"  # CN + US
    assert display_units(two_flags) == 4


def test_display_units_counts_hangul_syllables_as_wide() -> None:
    assert display_units("\uD55C\uAE00") == 4


def test_display_units_counts_emoji_as_two_units() -> None:
    assert display_units("\U0001F600") == 2
    assert display_units("\U0001F44D\U0001F3FD") == 2  # emoji + skin tone


def test_shorten_display_units_returns_original_when_within_limit() -> None:
    text = "Short text"
    assert shorten_display_units(text, max_units=64) == text


def test_shorten_display_units_adds_fullwidth_ellipsis() -> None:
    # 40 CJK characters = 80 display units; limit 64 -> prefix <= 62 + 2 suffix
    text = "\u4e00" * 40
    shortened = shorten_display_units(text, max_units=64)
    assert display_units(shortened) <= 64
    assert shortened.endswith("\u2026")


def test_shorten_display_units_respects_custom_max_units() -> None:
    text = "abcdefghij"
    # 10 ASCII units, limit 8 -> 6 prefix + ellipsis(2) = 8
    shortened = shorten_display_units(text, max_units=8)
    assert display_units(shortened) <= 8
    assert shortened.endswith("\u2026")


def test_shorten_display_units_preserves_grapheme_boundaries() -> None:
    # A cluster at the boundary must not be split.
    text = "abc\U0001F600de"
    shortened = shorten_display_units(text, max_units=5)
    # emoji is 2 units; prefix "abc" is 3 units, ellipsis is 2 -> total 5
    assert shortened == "abc\u2026"
    assert display_units(shortened) == 5




# Regression tests for unicode_text.py fixes


def test_hangul_syllable_plus_final_jamo_is_one_cluster() -> None:
    # 가 (U+AC00, LV syllable) + ㄱ (U+11A8, T jamo) must stay as one grapheme cluster.
    cluster = "각"
    assert display_units(cluster) == 2


def test_hangul_lvt_syllable_keeps_final_jamo() -> None:
    # 각 (U+AC01, LVT syllable) is internally treated as LVT, so adding another T is illegal
    # but the boundary after 각 is still a Hangul boundary.  Just verify it counts as one cluster.
    assert display_units("각") == 2


def test_thai_prepend_stays_with_following_consonant() -> None:
    # เ (U+0E40, Prepend) + ก (U+0E01, Thai consonant) is one visual cluster.
    assert display_units("เก") == 1


def test_mc_spacing_mark_attaches_to_previous_base() -> None:
    # क (U+0915, Devanagari consonant) + ◌़ (U+093C, SpacingMark) is one cluster.
    assert display_units("क़") == 1


def test_zwnj_attaches_to_previous_base_as_extend() -> None:
    # ZWNJ (U+200C) is Extend, so it joins the preceding base but does not bridge to the next base.
    assert display_units("a‌") == 1
    assert display_units("a‌b") == 2


def test_eaw_table_is_self_contained_no_runtime_fallback(monkeypatch: pytest.MonkeyPatch) -> None:
    # If _east_asian_width fell back to unicodedata.east_asian_width, this patch would break it.
    def broken_eaw(_codepoint: str) -> str:
        raise AssertionError("must not call runtime EAW")

    monkeypatch.setattr(unicodedata, "east_asian_width", broken_eaw)
    from skills.WPSComposer.scripts.longform import unicode_text as fresh

    assert fresh.display_units("中文") == 4
    assert fresh.display_units("Hello") == 5
    assert fresh._east_asian_width(0x4E00) == "W"
    assert fresh._east_asian_width(0x30000) == "W"
    assert fresh._east_asian_width(0xFF01) == "F"
    assert fresh._east_asian_width(0x2026) == "A"
    assert fresh._east_asian_width(0x0041) == "N"
