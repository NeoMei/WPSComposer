"""Unified colour model for WPSComposer.

Bridges human-readable ``#RRGGBB`` hex strings and COM BGR longs.
Also provides semantic colour-name resolution against a DesignPreset.
"""

from __future__ import annotations
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from .design_presets import DesignPreset


def hex_to_rgb_long(color):
    """Convert ``#RRGGBB`` string or int (already BGR-long) to COM colour long.

    Args:
        color: A ``#RRGGBB`` hex string or an integer already in BGR format.

    Returns:
        An integer suitable for COM ``.RGB`` properties.
    """
    if isinstance(color, str):
        c = color.lstrip("#")
        r = int(c[0:2], 16)
        g = int(c[2:4], 16)
        b = int(c[4:6], 16)
        return r + (g << 8) + (b << 16)
    return int(color)


def rgb_long_to_hex(bgr):
    """Convert a COM BGR long back to ``#RRGGBB`` string."""
    b = (bgr >> 16) & 0xFF
    g = (bgr >> 8) & 0xFF
    r = bgr & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


# ---------------------------------------------------------------------------
# Semantic colour-name → hex resolution against a DesignPreset
# ---------------------------------------------------------------------------

# Maps colour names used in layout templates to preset colour roles.
# Extend this when layouts introduce new semantic names.
_SEMANTIC_COLOR_MAP = {
    "primary":       "primary",
    "secondary":     "secondary",
    "accent":        "accent",
    "accent_light":  "accent",       # maps to accent (lighter variant)
    "dark":          "dark",
    "light":         "light",
    "light_text":    "light",
    "white":         "bg",           # assuming light bg for now
    "gray":          "dark",         # maps to dark (gray variant)
}


def resolve_color(name: str, preset: Optional["DesignPreset"] = None) -> str:
    """Resolve a semantic colour name to a ``#RRGGBB`` hex string.

    Args:
        name: Semantic name like ``"primary"``, ``"white"``, ``"accent_light"``,
              or a raw ``#RRGGBB`` hex string.
        preset: A DesignPreset to resolve against.  Required for semantic names.

    Returns:
        ``#RRGGBB`` hex string.

    Raises:
        ValueError: If the name is a semantic reference but no preset is given
                    or the name is unrecognised.
    """
    # Already a hex string
    if name.startswith("#"):
        return name

    # Resolve semantic name via preset
    role = _SEMANTIC_COLOR_MAP.get(name)
    if role is None:
        raise ValueError(
            f"Unknown colour name '{name}'. "
            f"Known names: {list(_SEMANTIC_COLOR_MAP)}"
        )
    if preset is None:
        raise ValueError(
            f"Colour name '{name}' requires a DesignPreset for resolution."
        )
    return preset.get_color(role)


def resolve_color_long(name: str, preset: Optional["DesignPreset"] = None) -> int:
    """Resolve a semantic colour name to a COM BGR long.

    Convenience wrapper combining ``resolve_color`` + ``hex_to_rgb_long``.
    """
    return hex_to_rgb_long(resolve_color(name, preset))


def hex_to_rgb_long_static(color):
    """Alias for hex_to_rgb_long (static-friendly name)."""
    return hex_to_rgb_long(color)
