'''Design presets — color/font/spacing schemes for quick document styling.

Integrated from harness-anything (yb2460/harness-anything).
Colors use #RRGGBB hex strings compatible with WPSComposer's _hex_to_rgb_long().

Provides 4 presets:
  - academic:    University defense / scientific slides
  - consultant:  Consulting / professional reports
  - business:    Corporate presentations
  - tech:        Dark-mode tech / minimalist
'''

from __future__ import annotations
from typing import Dict, Tuple


class DesignPreset:
    '''Unified design preset: colors, fonts, spacing, and style rules.

    Attributes:
        name: Human-readable preset name.
        colors: Dict of color role → '#RRGGBB' hex string.
        fonts: Dict of role → (family, size_pt, '#RRGGBB' color).
        spacing: Dict with margin, gap, card_padding, line_height (all pt values).
        rules: Dict of design constraints (visual_ratio, max_colors, etc.).
    '''

    def __init__(
        self,
        name: str,
        colors: Dict[str, str],
        fonts: Dict[str, Tuple[str, int, str]],
        spacing: Dict[str, float],
        rules: Dict[str, object],
    ):
        self.name = name
        self.colors = colors
        self.fonts = fonts
        self.spacing = spacing
        self.rules = rules

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "colors": self.colors,
            "fonts": self.fonts,
            "spacing": self.spacing,
            "rules": self.rules,
        }

    def get_color(self, role: str) -> str:
        '''Return #RRGGBB hex string for a color role (e.g. "primary").'''
        return self.colors.get(role, "#000000")

    def get_font(self, role: str) -> Tuple[str, int, str]:
        '''Return (family, size_pt, color_hex) for a font role (e.g. "title").'''
        return self.fonts.get(role, ("Arial", 12, "#000000"))


# ============================================================
# 4 preset themes
# ============================================================

PRESETS: Dict[str, DesignPreset] = {}

# --- 1. Academic (scientific-slides) ---
PRESETS["academic"] = DesignPreset(
    name="学术答辩",
    colors={
        "primary":   "#1A3C8B",
        "secondary": "#E67733",
        "accent":    "#188050",
        "dark":      "#222222",
        "light":     "#F5F8FC",
        "bg":        "#FFFFFF",
    },
    fonts={
        "title":     ("Arial", 40, "#1A3C8B"),
        "subtitle":  ("Arial", 22, "#646464"),
        "body":      ("Arial", 24, "#222222"),
        "caption":   ("Arial", 16, "#808080"),
        "chinese":   ("微软雅黑", 24, "#222222"),
    },
    spacing={
        "margin": 80,
        "gap": 24,
        "card_padding": 20,
        "line_height": 1.5,
    },
    rules={
        "visual_ratio": 0.65,
        "max_colors": 5,
        "colorblind_safe": True,
        "one_idea_per_slide": True,
        "white_space": 0.4,
        "max_bullets": 6,
        "prefer_sans_serif": True,
    },
)

# --- 2. Consultant ---
PRESETS["consultant"] = DesignPreset(
    name="咨询顾问",
    colors={
        "primary":   "#003366",
        "secondary": "#00A8E8",
        "accent":    "#FF8C00",
        "dark":      "#212121",
        "light":     "#F2F6FA",
        "bg":        "#FFFFFF",
    },
    fonts={
        "title":     ("Arial", 36, "#003366"),
        "subtitle":  ("Arial", 20, "#505050"),
        "body":      ("Arial", 18, "#212121"),
        "caption":   ("Arial", 14, "#8C8C8C"),
        "chinese":   ("微软雅黑", 18, "#212121"),
    },
    spacing={
        "margin": 60,
        "gap": 20,
        "card_padding": 16,
        "line_height": 1.4,
    },
    rules={
        "visual_ratio": 0.5,
        "max_colors": 4,
        "colorblind_safe": False,
        "one_idea_per_slide": True,
        "white_space": 0.4,
        "max_bullets": 5,
        "prefer_sans_serif": True,
        "grid_layout": True,
        "content_density": "medium",
    },
)

# --- 3. Business ---
PRESETS["business"] = DesignPreset(
    name="商务汇报",
    colors={
        "primary":   "#005294",
        "secondary": "#C82828",
        "accent":    "#2DA050",
        "dark":      "#2D2D30",
        "light":     "#F8F9FA",
        "bg":        "#FFFFFF",
    },
    fonts={
        "title":     ("Arial", 36, "#005294"),
        "subtitle":  ("Arial", 20, "#5A5A5A"),
        "body":      ("Arial", 18, "#2D2D30"),
        "caption":   ("Arial", 14, "#8C8C8C"),
        "chinese":   ("微软雅黑", 18, "#2D2D30"),
    },
    spacing={
        "margin": 70,
        "gap": 22,
        "card_padding": 18,
        "line_height": 1.4,
    },
    rules={
        "visual_ratio": 0.45,
        "max_colors": 4,
        "colorblind_safe": False,
        "one_idea_per_slide": True,
        "white_space": 0.35,
        "max_bullets": 6,
        "prefer_sans_serif": True,
        "content_density": "medium",
    },
)

# --- 4. Tech / Dark minimal ---
PRESETS["tech"] = DesignPreset(
    name="科技极简",
    colors={
        "primary":   "#0F1423",
        "secondary": "#00C8FF",
        "accent":    "#FF643C",
        "dark":      "#0F1423",
        "light":     "#F0F2F5",
        "bg":        "#0F1423",
    },
    fonts={
        "title":     ("Arial", 44, "#FFFFFF"),
        "subtitle":  ("Arial", 20, "#A0B4C8"),
        "body":      ("Arial", 20, "#FFFFFF"),
        "caption":   ("Arial", 14, "#788296"),
        "chinese":   ("微软雅黑", 20, "#FFFFFF"),
    },
    spacing={
        "margin": 100,
        "gap": 30,
        "card_padding": 24,
        "line_height": 1.6,
    },
    rules={
        "visual_ratio": 0.55,
        "max_colors": 3,
        "colorblind_safe": True,
        "one_idea_per_slide": True,
        "white_space": 0.5,
        "max_bullets": 4,
        "prefer_sans_serif": True,
        "dark_mode": True,
        "content_density": "low",
    },
)

# --- 5. Proposal (formal Chinese corporate/government docs) ---
PRESETS["proposal"] = DesignPreset(
    name="\u653f\u4f73\u65b9\u6848",  # 政企方案
    colors={
        "primary":   "#000000",          # black (formal docs use plain black)
        "secondary": "#333333",          # dark grey
        "accent":    "#1A3C8B",          # navy blue for tables
        "dark":      "#000000",
        "light":     "#F5F5F5",          # light grey for shading
        "bg":        "#FFFFFF",
    },
    fonts={
        "title":     ("\u9ed1\u4f53", 22, "#000000"),       # SimHei 22pt
        "subtitle":  ("\u9ed1\u4f53", 16, "#000000"),       # SimHei 16pt
        "body":      ("\u4eff\u5b8b", 12, "#000000"),       # FangSong 12pt
        "caption":   ("\u4eff\u5b8b", 9, "#666666"),        # FangSong 9pt
        "chinese":   ("\u4eff\u5b8b", 12, "#000000"),       # FangSong 12pt
    },
    spacing={
        "margin": 72,        # 1 inch
        "gap": 10,
        "card_padding": 12,
        "line_height": 1.5,
    },
    rules={
        "visual_ratio": 0.4,
        "max_colors": 3,
        "colorblind_safe": True,
        "one_idea_per_slide": True,
        "white_space": 0.3,
        "max_bullets": 6,
        "prefer_sans_serif": False,   # use serif/song style
        "content_density": "medium",
    },
)

# ============================================================
# Registration -- extendable preset system
# ============================================================

def register(preset: DesignPreset) -> DesignPreset:
    '''Register a new or replacement design preset.

    Args:
        preset: A DesignPreset instance.

    Returns:
        The preset (for use as a decorator).
    '''
    PRESETS[preset.name] = preset
    return preset


def unregister(name: str) -> None:
    '''Remove a preset by name.'''
    PRESETS.pop(name, None)


def register_from_dict(name: str, colors: dict, fonts: dict,
                       spacing: dict, rules: dict) -> DesignPreset:
    '''Create and register a preset from raw dicts.'''
    preset = DesignPreset(name, colors, fonts, spacing, rules)
    return register(preset)



def get_preset(name: str) -> DesignPreset:
    '''Return a preset by name.  Raises KeyError if not found.'''
    if name not in PRESETS:
        raise KeyError(
            f"Unknown preset '{name}'. Available: {', '.join(PRESETS)}"
        )
    return PRESETS[name]


def list_presets() -> list:
    '''Return list of available preset names.'''
    return list(PRESETS.keys())
