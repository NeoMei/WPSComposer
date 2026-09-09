"""Fixed Unicode 15.1.0 text helpers for the long-form semantic core.

These helpers intentionally do not depend on the Python runtime's Unicode
version.  East Asian Width, Script=Han, and grapheme-cluster tables are pinned
to the subsets required by M1; results are identical across Python versions.
"""

from __future__ import annotations

import unicodedata
from typing import Iterable, Iterator, Tuple


def normalize_visible_text(text: str) -> str:
    """Normalize visible text to Unicode NFC before it enters the semantic model."""
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    return unicodedata.normalize("NFC", text)


# ---------------------------------------------------------------------------
# Fixed Script=Han ranges (Unicode 15.1.0)
# ---------------------------------------------------------------------------
_HAN_RANGES: Tuple[Tuple[int, int], ...] = (
    (0x3400, 0x4DBF),    # CJK Unified Ideographs Extension A
    (0x4E00, 0x9FFF),    # CJK Unified Ideographs
    (0xF900, 0xFAFF),    # CJK Compatibility Ideographs
    (0x20000, 0x2A6DF),  # CJK Unified Ideographs Extension B
    (0x2A700, 0x2B73F),  # CJK Unified Ideographs Extension C
    (0x2B740, 0x2B81F),  # CJK Unified Ideographs Extension D
    (0x2B820, 0x2CEAF),  # CJK Unified Ideographs Extension E
    (0x2CEB0, 0x2EBEF),  # CJK Unified Ideographs Extension F
    (0x2F800, 0x2FA1F),  # CJK Compatibility Ideographs Supplement
    (0x30000, 0x3134F),  # CJK Unified Ideographs Extension G
)


def _in_ranges(codepoint: int, ranges: Iterable[Tuple[int, int]]) -> bool:
    """Binary-search-free range check; ranges are few and cold-path."""
    for start, end in ranges:
        if start <= codepoint <= end:
            return True
    return False


def contains_han(text: str) -> bool:
    """Return True if *text* contains at least one Unicode 15.1 Script=Han codepoint."""
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    for character in text:
        if _in_ranges(ord(character), _HAN_RANGES):
            return True
    return False


# ---------------------------------------------------------------------------
# Fixed East Asian Width table (Unicode 15.1.0)
# ---------------------------------------------------------------------------
# Self-contained W/F/A tables for the M1 required ranges.  This module never
# calls unicodedata.east_asian_width, so display-unit results are stable across
# Python runtime Unicode database versions.  CJK Extension G is included in _W.
_W_EAW_RANGES = (
    (0x1100, 0x115F),  # 96 codepoints
    (0x231A, 0x231B),  # 2 codepoints
    (0x2329, 0x232A),  # 2 codepoints
    (0x23E9, 0x23EC),  # 4 codepoints
    (0x23F0, 0x23F0),
    (0x23F3, 0x23F3),
    (0x25FD, 0x25FE),  # 2 codepoints
    (0x2614, 0x2615),  # 2 codepoints
    (0x2648, 0x2653),  # 12 codepoints
    (0x267F, 0x267F),
    (0x2693, 0x2693),
    (0x26A1, 0x26A1),
    (0x26AA, 0x26AB),  # 2 codepoints
    (0x26BD, 0x26BE),  # 2 codepoints
    (0x26C4, 0x26C5),  # 2 codepoints
    (0x26CE, 0x26CE),
    (0x26D4, 0x26D4),
    (0x26EA, 0x26EA),
    (0x26F2, 0x26F3),  # 2 codepoints
    (0x26F5, 0x26F5),
    (0x26FA, 0x26FA),
    (0x26FD, 0x26FD),
    (0x2705, 0x2705),
    (0x270A, 0x270B),  # 2 codepoints
    (0x2728, 0x2728),
    (0x274C, 0x274C),
    (0x274E, 0x274E),
    (0x2753, 0x2755),  # 3 codepoints
    (0x2757, 0x2757),
    (0x2795, 0x2797),  # 3 codepoints
    (0x27B0, 0x27B0),
    (0x27BF, 0x27BF),
    (0x2B1B, 0x2B1C),  # 2 codepoints
    (0x2B50, 0x2B50),
    (0x2B55, 0x2B55),
    (0x2E80, 0x2E99),  # 26 codepoints
    (0x2E9B, 0x2EF3),  # 89 codepoints
    (0x2F00, 0x2FD5),  # 214 codepoints
    (0x2FF0, 0x2FFB),  # 12 codepoints
    (0x3001, 0x303E),  # 62 codepoints
    (0x3041, 0x3096),  # 86 codepoints
    (0x3099, 0x30FF),  # 103 codepoints
    (0x3105, 0x312F),  # 43 codepoints
    (0x3131, 0x318E),  # 94 codepoints
    (0x3190, 0x31E3),  # 84 codepoints
    (0x31F0, 0x321E),  # 47 codepoints
    (0x3220, 0x3247),  # 40 codepoints
    (0x3250, 0x4DBF),  # 7024 codepoints
    (0x4E00, 0x9FFC),  # 20989 codepoints
    (0xA000, 0xA48C),  # 1165 codepoints
    (0xA490, 0xA4C6),  # 55 codepoints
    (0xA960, 0xA97C),  # 29 codepoints
    (0xAC00, 0xD7A3),  # 11172 codepoints
    (0xF900, 0xFA6D),  # 366 codepoints
    (0xFA70, 0xFAD9),  # 106 codepoints
    (0xFE10, 0xFE19),  # 10 codepoints
    (0xFE30, 0xFE52),  # 35 codepoints
    (0xFE54, 0xFE66),  # 19 codepoints
    (0xFE68, 0xFE6B),  # 4 codepoints
    (0x16FE0, 0x16FE4),  # 5 codepoints
    (0x16FF0, 0x16FF1),  # 2 codepoints
    (0x17000, 0x187F7),  # 6136 codepoints
    (0x18800, 0x18CD5),  # 1238 codepoints
    (0x18D00, 0x18D08),  # 9 codepoints
    (0x1B000, 0x1B11E),  # 287 codepoints
    (0x1B150, 0x1B152),  # 3 codepoints
    (0x1B164, 0x1B167),  # 4 codepoints
    (0x1B170, 0x1B2FB),  # 396 codepoints
    (0x1F004, 0x1F004),
    (0x1F0CF, 0x1F0CF),
    (0x1F18E, 0x1F18E),
    (0x1F191, 0x1F19A),  # 10 codepoints
    (0x1F200, 0x1F202),  # 3 codepoints
    (0x1F210, 0x1F23B),  # 44 codepoints
    (0x1F240, 0x1F248),  # 9 codepoints
    (0x1F250, 0x1F251),  # 2 codepoints
    (0x1F260, 0x1F265),  # 6 codepoints
    (0x1F300, 0x1F320),  # 33 codepoints
    (0x1F32D, 0x1F335),  # 9 codepoints
    (0x1F337, 0x1F37C),  # 70 codepoints
    (0x1F37E, 0x1F393),  # 22 codepoints
    (0x1F3A0, 0x1F3CA),  # 43 codepoints
    (0x1F3CF, 0x1F3D3),  # 5 codepoints
    (0x1F3E0, 0x1F3F0),  # 17 codepoints
    (0x1F3F4, 0x1F3F4),
    (0x1F3F8, 0x1F43E),  # 71 codepoints
    (0x1F440, 0x1F440),
    (0x1F442, 0x1F4FC),  # 187 codepoints
    (0x1F4FF, 0x1F53D),  # 63 codepoints
    (0x1F54B, 0x1F54E),  # 4 codepoints
    (0x1F550, 0x1F567),  # 24 codepoints
    (0x1F57A, 0x1F57A),
    (0x1F595, 0x1F596),  # 2 codepoints
    (0x1F5A4, 0x1F5A4),
    (0x1F5FB, 0x1F64F),  # 85 codepoints
    (0x1F680, 0x1F6C5),  # 70 codepoints
    (0x1F6CC, 0x1F6CC),
    (0x1F6D0, 0x1F6D2),  # 3 codepoints
    (0x1F6D5, 0x1F6D7),  # 3 codepoints
    (0x1F6EB, 0x1F6EC),  # 2 codepoints
    (0x1F6F4, 0x1F6FC),  # 9 codepoints
    (0x1F7E0, 0x1F7EB),  # 12 codepoints
    (0x1F90C, 0x1F93A),  # 47 codepoints
    (0x1F93C, 0x1F945),  # 10 codepoints
    (0x1F947, 0x1F978),  # 50 codepoints
    (0x1F97A, 0x1F9CB),  # 82 codepoints
    (0x1F9CD, 0x1F9FF),  # 51 codepoints
    (0x1FA70, 0x1FA74),  # 5 codepoints
    (0x1FA78, 0x1FA7A),  # 3 codepoints
    (0x1FA80, 0x1FA86),  # 7 codepoints
    (0x1FA90, 0x1FAA8),  # 25 codepoints
    (0x1FAB0, 0x1FAB6),  # 7 codepoints
    (0x1FAC0, 0x1FAC2),  # 3 codepoints
    (0x1FAD0, 0x1FAD6),  # 7 codepoints
    (0x20000, 0x2A6DD),  # 42718 codepoints
    (0x2A700, 0x2B734),  # 4149 codepoints
    (0x2B740, 0x2B81D),  # 222 codepoints
    (0x2B820, 0x2CEA1),  # 5762 codepoints
    (0x2CEB0, 0x2EBE0),  # 7473 codepoints
    (0x2F800, 0x2FA1D),  # 542 codepoints
    (0x30000, 0x3134F),  # 4944 codepoints
)

_F_EAW_RANGES = (
    (0x3000, 0x3000),
    (0xFF01, 0xFF60),  # 96 codepoints
    (0xFFE0, 0xFFE6),  # 7 codepoints
)

_A_EAW_RANGES = (
    (0x00A1, 0x00A1),
    (0x00A4, 0x00A4),
    (0x00A7, 0x00A8),  # 2 codepoints
    (0x00AA, 0x00AA),
    (0x00AD, 0x00AE),  # 2 codepoints
    (0x00B0, 0x00B4),  # 5 codepoints
    (0x00B6, 0x00BA),  # 5 codepoints
    (0x00BC, 0x00BF),  # 4 codepoints
    (0x00C6, 0x00C6),
    (0x00D0, 0x00D0),
    (0x00D7, 0x00D8),  # 2 codepoints
    (0x00DE, 0x00E1),  # 4 codepoints
    (0x00E6, 0x00E6),
    (0x00E8, 0x00EA),  # 3 codepoints
    (0x00EC, 0x00ED),  # 2 codepoints
    (0x00F0, 0x00F0),
    (0x00F2, 0x00F3),  # 2 codepoints
    (0x00F7, 0x00FA),  # 4 codepoints
    (0x00FC, 0x00FC),
    (0x00FE, 0x00FE),
    (0x0101, 0x0101),
    (0x0111, 0x0111),
    (0x0113, 0x0113),
    (0x011B, 0x011B),
    (0x0126, 0x0127),  # 2 codepoints
    (0x012B, 0x012B),
    (0x0131, 0x0133),  # 3 codepoints
    (0x0138, 0x0138),
    (0x013F, 0x0142),  # 4 codepoints
    (0x0144, 0x0144),
    (0x0148, 0x014B),  # 4 codepoints
    (0x014D, 0x014D),
    (0x0152, 0x0153),  # 2 codepoints
    (0x0166, 0x0167),  # 2 codepoints
    (0x016B, 0x016B),
    (0x01CE, 0x01CE),
    (0x01D0, 0x01D0),
    (0x01D2, 0x01D2),
    (0x01D4, 0x01D4),
    (0x01D6, 0x01D6),
    (0x01D8, 0x01D8),
    (0x01DA, 0x01DA),
    (0x01DC, 0x01DC),
    (0x0251, 0x0251),
    (0x0261, 0x0261),
    (0x02C4, 0x02C4),
    (0x02C7, 0x02C7),
    (0x02C9, 0x02CB),  # 3 codepoints
    (0x02CD, 0x02CD),
    (0x02D0, 0x02D0),
    (0x02D8, 0x02DB),  # 4 codepoints
    (0x02DD, 0x02DD),
    (0x02DF, 0x02DF),
    (0x0300, 0x036F),  # 112 codepoints
    (0x0391, 0x03A1),  # 17 codepoints
    (0x03A3, 0x03A9),  # 7 codepoints
    (0x03B1, 0x03C1),  # 17 codepoints
    (0x03C3, 0x03C9),  # 7 codepoints
    (0x0401, 0x0401),
    (0x0410, 0x044F),  # 64 codepoints
    (0x0451, 0x0451),
    (0x2010, 0x2010),
    (0x2013, 0x2016),  # 4 codepoints
    (0x2018, 0x2019),  # 2 codepoints
    (0x201C, 0x201D),  # 2 codepoints
    (0x2020, 0x2022),  # 3 codepoints
    (0x2024, 0x2027),  # 4 codepoints
    (0x2030, 0x2030),
    (0x2032, 0x2033),  # 2 codepoints
    (0x2035, 0x2035),
    (0x203B, 0x203B),
    (0x203E, 0x203E),
    (0x2074, 0x2074),
    (0x207F, 0x207F),
    (0x2081, 0x2084),  # 4 codepoints
    (0x20AC, 0x20AC),
    (0x2103, 0x2103),
    (0x2105, 0x2105),
    (0x2109, 0x2109),
    (0x2113, 0x2113),
    (0x2116, 0x2116),
    (0x2121, 0x2122),  # 2 codepoints
    (0x2126, 0x2126),
    (0x212B, 0x212B),
    (0x2153, 0x2154),  # 2 codepoints
    (0x215B, 0x215E),  # 4 codepoints
    (0x2160, 0x216B),  # 12 codepoints
    (0x2170, 0x2179),  # 10 codepoints
    (0x2189, 0x2189),
    (0x2190, 0x2199),  # 10 codepoints
    (0x21B8, 0x21B9),  # 2 codepoints
    (0x21D2, 0x21D2),
    (0x21D4, 0x21D4),
    (0x21E7, 0x21E7),
    (0x2200, 0x2200),
    (0x2202, 0x2203),  # 2 codepoints
    (0x2207, 0x2208),  # 2 codepoints
    (0x220B, 0x220B),
    (0x220F, 0x220F),
    (0x2211, 0x2211),
    (0x2215, 0x2215),
    (0x221A, 0x221A),
    (0x221D, 0x2220),  # 4 codepoints
    (0x2223, 0x2223),
    (0x2225, 0x2225),
    (0x2227, 0x222C),  # 6 codepoints
    (0x222E, 0x222E),
    (0x2234, 0x2237),  # 4 codepoints
    (0x223C, 0x223D),  # 2 codepoints
    (0x2248, 0x2248),
    (0x224C, 0x224C),
    (0x2252, 0x2252),
    (0x2260, 0x2261),  # 2 codepoints
    (0x2264, 0x2267),  # 4 codepoints
    (0x226A, 0x226B),  # 2 codepoints
    (0x226E, 0x226F),  # 2 codepoints
    (0x2282, 0x2283),  # 2 codepoints
    (0x2286, 0x2287),  # 2 codepoints
    (0x2295, 0x2295),
    (0x2299, 0x2299),
    (0x22A5, 0x22A5),
    (0x22BF, 0x22BF),
    (0x2312, 0x2312),
    (0x2460, 0x24E9),  # 138 codepoints
    (0x24EB, 0x254B),  # 97 codepoints
    (0x2550, 0x2573),  # 36 codepoints
    (0x2580, 0x258F),  # 16 codepoints
    (0x2592, 0x2595),  # 4 codepoints
    (0x25A0, 0x25A1),  # 2 codepoints
    (0x25A3, 0x25A9),  # 7 codepoints
    (0x25B2, 0x25B3),  # 2 codepoints
    (0x25B6, 0x25B7),  # 2 codepoints
    (0x25BC, 0x25BD),  # 2 codepoints
    (0x25C0, 0x25C1),  # 2 codepoints
    (0x25C6, 0x25C8),  # 3 codepoints
    (0x25CB, 0x25CB),
    (0x25CE, 0x25D1),  # 4 codepoints
    (0x25E2, 0x25E5),  # 4 codepoints
    (0x25EF, 0x25EF),
    (0x2605, 0x2606),  # 2 codepoints
    (0x2609, 0x2609),
    (0x260E, 0x260F),  # 2 codepoints
    (0x261C, 0x261C),
    (0x261E, 0x261E),
    (0x2640, 0x2640),
    (0x2642, 0x2642),
    (0x2660, 0x2661),  # 2 codepoints
    (0x2663, 0x2665),  # 3 codepoints
    (0x2667, 0x266A),  # 4 codepoints
    (0x266C, 0x266D),  # 2 codepoints
    (0x266F, 0x266F),
    (0x269E, 0x269F),  # 2 codepoints
    (0x26BF, 0x26BF),
    (0x26C6, 0x26CD),  # 8 codepoints
    (0x26CF, 0x26D3),  # 5 codepoints
    (0x26D5, 0x26E1),  # 13 codepoints
    (0x26E3, 0x26E3),
    (0x26E8, 0x26E9),  # 2 codepoints
    (0x26EB, 0x26F1),  # 7 codepoints
    (0x26F4, 0x26F4),
    (0x26F6, 0x26F9),  # 4 codepoints
    (0x26FB, 0x26FC),  # 2 codepoints
    (0x26FE, 0x26FF),  # 2 codepoints
    (0x273D, 0x273D),
    (0x2776, 0x277F),  # 10 codepoints
    (0x2B56, 0x2B59),  # 4 codepoints
    (0x3248, 0x324F),  # 8 codepoints
    (0xE000, 0xF8FF),  # 6400 codepoints
    (0xFE00, 0xFE0F),  # 16 codepoints
    (0xFFFD, 0xFFFD),
    (0x1F100, 0x1F10A),  # 11 codepoints
    (0x1F110, 0x1F12D),  # 30 codepoints
    (0x1F130, 0x1F169),  # 58 codepoints
    (0x1F170, 0x1F18D),  # 30 codepoints
    (0x1F18F, 0x1F190),  # 2 codepoints
    (0x1F19B, 0x1F1AC),  # 18 codepoints
    (0xE0100, 0xE01EF),  # 240 codepoints
    (0xF0000, 0xFFFFD),  # 65534 codepoints
    (0x100000, 0x10FFFD),  # 65534 codepoints
)


def _east_asian_width(codepoint: int) -> str:
    """Return the East Asian Width property class for a scalar codepoint."""
    if _in_ranges(codepoint, _W_EAW_RANGES):
        return "W"
    if _in_ranges(codepoint, _F_EAW_RANGES):
        return "F"
    if _in_ranges(codepoint, _A_EAW_RANGES):
        return "A"
    return "N"


# ---------------------------------------------------------------------------
# Fixed emoji / Extended_Pictographic ranges (Unicode 15.1.0 subset)
# ---------------------------------------------------------------------------
_EMOJI_RANGES: Tuple[Tuple[int, int], ...] = (
    (0x231A, 0x231B),
    (0x23E9, 0x23EC),
    (0x23F0, 0x23F0),
    (0x23F3, 0x23F3),
    (0x25FD, 0x25FE),
    (0x2614, 0x2615),
    (0x2648, 0x2653),
    (0x267F, 0x267F),
    (0x2693, 0x2693),
    (0x26A1, 0x26A1),
    (0x26AA, 0x26AB),
    (0x26BD, 0x26BE),
    (0x26C4, 0x26C5),
    (0x26CE, 0x26CE),
    (0x26D4, 0x26D4),
    (0x26EA, 0x26EA),
    (0x26F2, 0x26F3),
    (0x26F5, 0x26F5),
    (0x26F7, 0x26FA),
    (0x26FD, 0x26FD),
    (0x2705, 0x2705),
    (0x2728, 0x2728),
    (0x274C, 0x274C),
    (0x274E, 0x274E),
    (0x2753, 0x2755),
    (0x2795, 0x2797),
    (0x27B0, 0x27B0),
    (0x27BF, 0x27BF),
    (0x2B50, 0x2B55),
    (0x1F004, 0x1F004),
    (0x1F0CF, 0x1F0CF),
    (0x1F1E6, 0x1F1FF),
    (0x1F300, 0x1F5FF),
    (0x1F600, 0x1F64F),
    (0x1F680, 0x1F6FF),
    (0x1F700, 0x1F77F),
    (0x1F780, 0x1F7FF),
    (0x1F800, 0x1F8FF),
    (0x1F900, 0x1F9FF),
    (0x1FA00, 0x1FAFF),
)


def _is_emoji(codepoint: int) -> bool:
    return _in_ranges(codepoint, _EMOJI_RANGES)


# Thai/Lao leading vowels and related prependers (UAX #29 Prepend / GB9b).
_PREPEND_RANGES = (
    (0x0E40, 0x0E44),  # Thai sara e / sara ae / sara o / sara ai / sara am
    (0x0EC0, 0x0EC4),  # Lao leading vowels
)


# ---------------------------------------------------------------------------
# Extended grapheme cluster boundaries (UAX #29, Unicode 15.1.0 subset)
# ---------------------------------------------------------------------------
_HANGUL_SBASE = 0xAC00
_HANGUL_LBASE = 0x1100
_HANGUL_VBASE = 0x1161
_HANGUL_TBASE = 0x11A7
_HANGUL_LCOUNT = 19
_HANGUL_VCOUNT = 21
_HANGUL_TCOUNT = 28
_HANGUL_NCOUNT = _HANGUL_VCOUNT * _HANGUL_TCOUNT  # 588
_HANGUL_SCOUNT = _HANGUL_LCOUNT * _HANGUL_NCOUNT  # 11172


def _hangul_class(codepoint: int) -> str | None:
    """Classify Hangul Jamo/Syllable for grapheme-cluster rules (UAX #29)."""
    if (0x1100 <= codepoint <= 0x115F) or (0xA960 <= codepoint <= 0xA97C):
        return "L"
    if (0x1160 <= codepoint <= 0x11A2) or (0xD7B0 <= codepoint <= 0xD7C6):
        return "V"
    if (0x11A8 <= codepoint <= 0x11F9) or (0xD7CB <= codepoint <= 0xD7FB):
        return "T"
    if _HANGUL_SBASE <= codepoint < _HANGUL_SBASE + _HANGUL_SCOUNT:
        # UAX #29 derives LV/LVT arithmetically from the syllable index.
        s_index = codepoint - _HANGUL_SBASE
        t_index = s_index % _HANGUL_TCOUNT
        return "LVT" if t_index else "LV"
    return None


def _grapheme_category(codepoint: int) -> str:
    """Classify a codepoint for grapheme-boundary decisions."""
    if codepoint == 0x000D:
        return "CR"
    if codepoint == 0x000A:
        return "LF"

    category = unicodedata.category(chr(codepoint))
    if category == "Cc":
        return "Control"
    # ZWNJ and ZWJ are Extend per UAX #29, not Control.
    if codepoint == 0x200C:
        return "Extend"
    if codepoint == 0x200D:
        return "ZWJ"

    hangul = _hangul_class(codepoint)
    if hangul is not None:
        return f"Hangul_{hangul}"

    # Thai/Lao leading vowels keep the next base in the same cluster (GB9b).
    if _in_ranges(codepoint, _PREPEND_RANGES):
        return "Prepend"

    # SpacingMark (Mc) attaches to the previous base (GB9a); Mn/Me extend.
    if category == "Mc":
        return "SpacingMark"
    if category in ("Mn", "Me"):
        return "Extend"
    # Skin-tone modifiers are Extend even though Python 3.9 reports them as Sk.
    if 0x1F3FB <= codepoint <= 0x1F3FF:
        return "Extend"

    if 0x1F1E6 <= codepoint <= 0x1F1FF:
        return "RegionalIndicator"
    if _is_emoji(codepoint):
        return "Emoji"
    return "Other"


def _grapheme_clusters(text: str) -> Iterator[str]:
    """Yield extended grapheme clusters of *text* in order."""
    if not text:
        return

    cluster_start = 0
    previous_category: str | None = None
    regional_indicator_open = False

    for index, character in enumerate(text):
        if index == 0:
            previous_category = _grapheme_category(ord(character))
            regional_indicator_open = previous_category == "RegionalIndicator"
            continue

        category = _grapheme_category(ord(character))
        break_cluster = True

        # GB3: CR x LF
        if previous_category == "CR" and category == "LF":
            break_cluster = False
        # GB4/GB5: controls break on either side (except CR x LF above).
        elif previous_category in ("Control", "CR", "LF"):
            break_cluster = True
        elif category in ("Control", "CR", "LF"):
            break_cluster = True
        # GB6-GB8: Hangul syllable sequences.
        elif previous_category == "Hangul_L" and category in (
            "Hangul_L",
            "Hangul_V",
            "Hangul_LV",
            "Hangul_LVT",
        ):
            break_cluster = False
        elif previous_category in ("Hangul_LV", "Hangul_V") and category in (
            "Hangul_V",
            "Hangul_T",
        ):
            break_cluster = False
        elif previous_category in ("Hangul_LVT", "Hangul_T") and category == "Hangul_T":
            break_cluster = False
        # GB9: do not break before Extend or ZWJ.
        elif category in ("Extend", "ZWJ"):
            break_cluster = False
        # GB9a: do not break before SpacingMark.
        elif category == "SpacingMark":
            break_cluster = False
        # GB9b: do not break after Prepend.
        elif previous_category == "Prepend":
            break_cluster = False
        # GB11: emoji ZWJ sequences (Emoji x ZWJ x Emoji).
        elif previous_category == "Emoji" and category == "ZWJ":
            break_cluster = False
        elif previous_category == "ZWJ" and category == "Emoji":
            break_cluster = False
        # GB12/GB13: pair regional indicators.
        elif previous_category == "RegionalIndicator" and category == "RegionalIndicator" and regional_indicator_open:
            break_cluster = False
        # Everything else breaks.

        if break_cluster:
            yield text[cluster_start:index]
            cluster_start = index
            regional_indicator_open = category == "RegionalIndicator"
        else:
            # Close an open RI when we pair it.
            if previous_category == "RegionalIndicator" and category == "RegionalIndicator":
                regional_indicator_open = False
            elif category == "RegionalIndicator":
                regional_indicator_open = True

        previous_category = category

    yield text[cluster_start:]


# ---------------------------------------------------------------------------
# Display-unit counting and header shortening
# ---------------------------------------------------------------------------
_FULLWIDTH_ELLIPSIS = "…"


def _cluster_display_units(cluster: str) -> int:
    """Return display units for a single grapheme cluster."""
    for character in cluster:
        codepoint = ord(character)
        if _is_emoji(codepoint):
            return 2
        if _east_asian_width(codepoint) in ("W", "F", "A"):
            return 2
    return 1


def display_units(text: str) -> int:
    """Return the number of display units in *text*.

    Each extended grapheme cluster contributes 2 units when it contains a
    wide/fullwidth/ambiguous East Asian Width character or an emoji character,
    otherwise 1 unit.
    """
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    text = normalize_visible_text(text)
    return sum(_cluster_display_units(cluster) for cluster in _grapheme_clusters(text))


def shorten_display_units(text: str, max_units: int = 64) -> str:
    """Shorten *text* to no more than *max_units* display units.

    A fullwidth ellipsis (U+2026) is appended at the last complete grapheme
    boundary; the ellipsis itself consumes 2 display units.
    """
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    if max_units <= 0:
        return ""

    text = normalize_visible_text(text)
    if display_units(text) <= max_units:
        return text

    suffix = _FULLWIDTH_ELLIPSIS
    suffix_units = _cluster_display_units(suffix)
    available = max_units - suffix_units
    if available <= 0:
        return ""

    used = 0
    end_index = 0
    for cluster in _grapheme_clusters(text):
        units = _cluster_display_units(cluster)
        if used + units > available:
            break
        used += units
        end_index += len(cluster)

    return text[:end_index] + suffix


__all__ = [
    "contains_han",
    "display_units",
    "normalize_visible_text",
    "shorten_display_units",
]