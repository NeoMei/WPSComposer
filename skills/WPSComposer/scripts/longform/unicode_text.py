"""Fixed Unicode 15.1.0 text helpers for the long-form semantic core.

These helpers intentionally do not depend on the Python runtime's Unicode
version.  East Asian Width and Script=Han tables are pinned to the subsets
required by M1; grapheme cluster boundaries follow UAX #29 extended clusters
closely enough for deterministic page-header shortening and heading-script
selection.
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
# Fixed East Asian Width overrides for ranges absent or wrong in Python 3.9
# ---------------------------------------------------------------------------
_EAW_OVERRIDES: dict[str, Tuple[Tuple[int, int], ...]] = {
    # Unicode 15.1 CJK Extension G is unassigned in Python 3.9's Unicode 13.0.
    "W": ((0x30000, 0x3134F),),
}


def _east_asian_width(codepoint: int) -> str:
    """Return the East Asian Width property class for a scalar codepoint."""
    for width, ranges in _EAW_OVERRIDES.items():
        if _in_ranges(codepoint, ranges):
            return width
    return unicodedata.east_asian_width(chr(codepoint))


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


# ---------------------------------------------------------------------------
# Extended grapheme cluster boundaries (UAX #29, Unicode 15.1.0 subset)
# ---------------------------------------------------------------------------
def _hangul_class(codepoint: int) -> str | None:
    """Classify Hangul Jamo/Syllable for grapheme-cluster rules."""
    if (0x1100 <= codepoint <= 0x115F) or (0xA960 <= codepoint <= 0xA97C):
        return "L"
    if (0x1160 <= codepoint <= 0x11A2) or (0xD7B0 <= codepoint <= 0xD7C6):
        return "V"
    if (0x11A8 <= codepoint <= 0x11F9) or (0xD7CB <= codepoint <= 0xD7FB):
        return "T"
    if 0xAC00 <= codepoint <= 0xD7A3:
        decomposition = unicodedata.decomposition(chr(codepoint))
        if decomposition:
            parts = decomposition.split()
            if len(parts) == 2:
                return "LV"
            if len(parts) == 3:
                return "LVT"
    return None


def _grapheme_category(codepoint: int) -> str:
    """Classify a codepoint for grapheme-boundary decisions."""
    if codepoint == 0x000D:
        return "CR"
    if codepoint == 0x000A:
        return "LF"
    # Control characters.  ZWNJ is treated as a control per UAX #29 GCB.
    category = unicodedata.category(chr(codepoint))
    if category == "Cc" or codepoint == 0x200C:
        return "Control"
    if codepoint == 0x200D:
        return "ZWJ"

    hangul = _hangul_class(codepoint)
    if hangul is not None:
        return f"Hangul_{hangul}"

    # Extend / spacing mark.  Include skin-tone modifiers explicitly because
    # Python 3.9 reports them as 'Sk' while UAX #29 treats them as Extend.
    if category in ("Mn", "Me", "Mc"):
        return "Extend"
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
