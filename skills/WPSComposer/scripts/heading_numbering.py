"""Create heading_numbering.py — automatic heading numbering system.

Handles three concerns:
1. Detects if user already wrote manual numbering (1.1, 一、, etc.)
2. Generates automatic multi-level numbering when needed
3. Provides Chinese and Arabic numbering schemes
"""

from __future__ import annotations
import re
from typing import List, Optional


# ---------------------------------------------------------------------------
# Manual numbering detection
# ---------------------------------------------------------------------------

# Patterns for manual numbering at start of heading text
_MANUAL_PATTERNS = [
    re.compile(r"^\d+\.\d+(\.\d+)*\s+"),     # 1.1 / 1.2.3
    re.compile(r"^\d+\.\s+"),                  # 1.
    re.compile(r"^\d+\s+"),                     # 1
    re.compile(r"^[\u4e00-\u9fff]{1,2}[\u3001\uff0c\u3002]\s*"),  # 一、 二，
    re.compile(r"^[\uff08\u3010]\s*[\u4e00-\u9fff]{1,3}\s*[\uff09\u3011]\s*"),  # （一） 【一】
    re.compile(r"^Chapter\s+\d+", re.IGNORECASE),
    re.compile(r"^Section\s+\d+", re.IGNORECASE),
]


def has_manual_numbering(text: str) -> bool:
    """Check if heading text already contains manual numbering prefix.

    Returns True for:
        "1.1 第一节", "一、概述", "（一）背景", "Chapter 1"
    Returns False for:
        "第一节", "概述", "系统设计"
    """
    for pat in _MANUAL_PATTERNS:
        if pat.match(text):
            return True
    return False


def strip_manual_numbering(text: str) -> str:
    """Remove manual numbering prefix, return clean heading text."""
    for pat in _MANUAL_PATTERNS:
        m = pat.match(text)
        if m:
            return text[m.end():].strip()
    return text


# ---------------------------------------------------------------------------
# Automatic numbering generator
# ---------------------------------------------------------------------------

_CN_NUMS = "\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341"  # 一二三四五六七八九十


def _num_to_cn(n: int) -> str:
    """Convert number to Chinese ordinal (1->一, 11->十一, 20->二十)."""
    if n <= 0:
        return str(n)
    if n <= 10:
        return _CN_NUMS[n - 1]
    if n < 20:
        return "\u5341" + _CN_NUMS[n - 11]  # 十一, 十二
    if n == 20:
        return "\u4e8c\u5341"  # 二十
    if n < 100:
        tens = n // 10
        ones = n % 10
        result = _CN_NUMS[tens - 1] + "\u5341"
        if ones > 0:
            result += _CN_NUMS[ones - 1]
        return result
    return str(n)


class NumberingState:
    """Track multi-level heading numbering across a document.

    Usage:
        ns = NumberingState(scheme="arabic")  # or "chinese"
        ns.advance(level=1)  # -> returns "1"
        ns.advance(level=2)  # -> returns "1.1"
        ns.advance(level=2)  # -> returns "1.2"
        ns.advance(level=1)  # -> returns "2"
    """

    def __init__(self, scheme: str = "arabic"):
        """Initialize numbering state.

        Args:
            scheme: "arabic" (1, 1.1, 1.1.1) or "chinese"
                    (第一章, 第一节, 一、).
        """
        self.scheme = scheme
        self.counters = [0, 0, 0, 0, 0, 0]  # levels 1-6

    def advance(self, level: int) -> str:
        """Advance the counter at *level* and return the number string.

        Resets all deeper levels to 0.

        Args:
            level: Heading level (1-6).

        Returns:
            Number string like "1", "1.2", "1.2.3" (arabic)
            or "\u7b2c\u4e00\u7ae0", "\u7b2c\u4e00\u8282" (chinese).
        """
        if level < 1 or level > 6:
            level = max(1, min(6, level))

        # Increment current level
        self.counters[level - 1] += 1
        # Reset deeper levels
        for i in range(level, 6):
            self.counters[i] = 0

        if self.scheme == "chinese":
            return self._format_chinese(level)
        else:
            return self._format_arabic(level)

    def _format_arabic(self, level: int) -> str:
        """Format as dotted decimal: 1, 1.1, 1.2.3."""
        parts = [str(self.counters[i]) for i in range(level) if self.counters[i] > 0]
        return ".".join(parts)

    def _format_chinese(self, level: int) -> str:
        """Format as Chinese: \u7b2c\u4e00\u7ae0 / \u7b2c\u4e00\u8282 / \u4e00\u3001."""
        n = self.counters[level - 1]
        if level == 1:
            return f"\u7b2c{_num_to_cn(n)}\u7ae0"       # 第N章
        elif level == 2:
            return f"\u7b2c{_num_to_cn(n)}\u8282"       # 第N节
        elif level == 3:
            return f"{_num_to_cn(n)}\u3001"              # 一、
        elif level == 4:
            return f"\uff08{_num_to_cn(n)}\uff09"        # （一）
        elif level == 5:
            return f"{n}."                               # 1.
        else:
            return f"({n})"                              # (1)


def detect_numbering_scheme(sections: list) -> str:
    """Heuristically detect whether to use Chinese or Arabic numbering.

    Scans heading texts for Chinese numbering hints.
    """
    cn_hints = 0
    total = 0
    for sec in sections:
        if sec.has_heading:
            total += 1
            text = sec.heading
            # Check for Chinese characters in heading
            if any("\u4e00" <= c <= "\u9fff" for c in text):
                cn_hints += 1
    # If most headings contain Chinese, use Chinese scheme
    return "chinese" if cn_hints > total * 0.5 and total > 0 else "arabic"
