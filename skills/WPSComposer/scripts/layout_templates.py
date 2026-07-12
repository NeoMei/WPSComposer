'''Layout templates — 14-slide narrative structure for presentations.

Integrated from harness-anything (yb2460/harness-anything).
Provides standard PPT slide layouts and talk-type presets (conference,
business, defense, school).
'''

from __future__ import annotations
from typing import Dict, List, Optional


class LayoutTemplate:
    '''Single-slide layout template.

    Attributes:
        name: Layout name (Chinese).
        category: cover / content / column / grid / image / close.
        description: What this layout is used for.
        elements: List of dict element descriptors with type, x, y, w, h,
                  and type-specific props (text, fs, color, align, etc.).
        structural_rules: Optional dict of layout-specific rules.
    '''

    def __init__(
        self,
        name: str,
        category: str,
        description: str,
        elements: List[dict],
        structural_rules: Optional[dict] = None,
    ):
        self.name = name
        self.category = category
        self.description = description
        self.elements = elements
        self.structural_rules = structural_rules or {}

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "category": self.category,
            "description": self.description,
            "elements": self.elements,
        }


# ============================================================
# 14 standard slide layouts
# ============================================================

LAYOUTS: Dict[str, LayoutTemplate] = {}

# --- Cover ---
LAYOUTS["cover"] = LayoutTemplate(
    name="封面", category="cover",
    description="深色全屏 + 72pt大字标题 + 装饰线 + 副标题/日期",
    elements=[
        {"type": "bg", "color": "primary"},
        {"type": "line", "x": 0, "y": 0, "w": 960, "h": 18, "color": "secondary"},
        {"type": "text", "x": 0, "y": 100, "w": 960, "h": 150, "text": "{title}", "fs": 72, "color": "white", "bold": True, "align": 2},
        {"type": "line", "x": 280, "y": 270, "w": 400, "h": 5, "color": "white"},
        {"type": "text", "x": 0, "y": 365, "w": 960, "h": 60, "text": "{subtitle}", "fs": 32, "color": "accent_light", "bold": True, "align": 2},
        {"type": "line", "x": 0, "y": 530, "w": 960, "h": 8, "color": "secondary"},
    ],
)

# --- TOC ---
LAYOUTS["toc"] = LayoutTemplate(
    name="目录", category="content",
    description="左侧蓝色装饰条 + 编号圆角方块 + 6项导航",
    elements=[
        {"type": "line", "x": 0, "y": 0, "w": 22, "h": 540, "color": "primary"},
        {"type": "text", "x": 60, "y": 25, "w": 400, "h": 65, "text": "目  录", "fs": 50, "color": "primary", "bold": True},
        {"type": "line", "x": 60, "y": 100, "w": 120, "h": 5, "color": "secondary"},
        {"type": "grid_items", "x": 60, "y": 140, "cols": 1, "rows": 6, "gap": 63, "item_w": 840, "item_h": 56,
         "fields": ["{num}", "{title}", "{desc}"],
         "styles": [{"fs": 24, "color": "white", "bold": True, "bg": "primary", "w": 56},
                    {"fs": 28, "color": "dark", "bold": True},
                    {"fs": 14, "color": "gray"}]},
    ],
)

# --- Overview ---
LAYOUTS["overview"] = LayoutTemplate(
    name="概览", category="content",
    description="顶部深色横幅 + 信息卡片矩阵 + 侧边荣誉栏",
    elements=[
        {"type": "line", "x": 0, "y": 0, "w": 960, "h": 100, "color": "primary"},
        {"type": "text", "x": 0, "y": 15, "w": 960, "h": 50, "text": "{title}", "fs": 42, "color": "white", "bold": True, "align": 2},
        {"type": "grid_cards", "x": 35, "y": 120, "cols": 2, "rows": 4, "gap_x": 20, "gap_y": 88, "item_w": 430, "item_h": 80,
         "fields": ["{label}", "{value}"],
         "styles": [{"fs": 16, "color": "white", "bold": True, "bg": "primary", "h": 38},
                    {"fs": 19, "color": "dark", "bold": True}]},
        {"type": "sidebar", "x": 720, "y": 120, "w": 200, "h": 360, "color": "light", "text": "{sidebar}"},
    ],
)

# --- Timeline ---
LAYOUTS["timeline"] = LayoutTemplate(
    name="时间轴", category="column",
    description="左侧圆点+竖线+右侧事件，适合历史/发展内容",
    elements=[
        {"type": "text", "x": 35, "y": 30, "w": 800, "h": 50, "text": "{title}", "fs": 38, "color": "primary", "bold": True},
        {"type": "line", "x": 35, "y": 85, "w": 150, "h": 3, "color": "secondary"},
        {"type": "timeline_items", "x": 40, "start_y": 120, "count": 6, "gap": 55,
         "fields": ["{date}", "{event}"],
         "styles": [{"fs": 16, "color": "primary", "bold": True, "w": 110},
                    {"fs": 16, "color": "dark", "w": 550}]},
        {"type": "sidebar", "x": 650, "y": 120, "w": 280, "h": 360, "color": "light", "text": "{sidebar}"},
    ],
)

# --- Grid Cards ---
LAYOUTS["grid_cards"] = LayoutTemplate(
    name="卡片网格", category="grid",
    description="2-4列卡片网格，适合展示并行信息（人物/地标/产品/特征）",
    elements=[
        {"type": "text", "x": 0, "y": 15, "w": 960, "h": 55, "text": "{title}", "fs": 44, "color": "primary", "bold": True, "align": 2},
        {"type": "line", "x": 350, "y": 72, "w": 260, "h": 3, "color": "secondary"},
        {"type": "grid_cards", "x": 60, "y": 95, "cols": 3, "rows": 2, "gap_x": 20, "gap_y": 20, "item_w": 260, "item_h": 200,
         "fields": ["{card_title}", "{card_content}"],
         "styles": [{"fs": 18, "color": "primary", "bold": True, "align": 2, "h": 40},
                    {"fs": 13, "color": "dark", "w": 240, "h": 150}]},
    ],
)

# --- Quadrant (2x2) ---
LAYOUTS["quadrant"] = LayoutTemplate(
    name="四象限", category="grid",
    description="2x2 四象限布局，适合对比/分类/SWOT分析",
    elements=[
        {"type": "text", "x": 30, "y": 15, "w": 500, "h": 50, "text": "{title}", "fs": 40, "color": "primary", "bold": True},
        {"type": "line", "x": 30, "y": 70, "w": 180, "h": 4, "color": "secondary"},
        {"type": "quadrant", "x": 30, "y": 90, "w": 900, "h": 440, "cols": 2, "rows": 2,
         "fields": ["{q_title}", "{q_content}"],
         "styles": [{"fs": 20, "color": "white", "bold": True, "bg": "{q_color}", "h": 45},
                    {"fs": 14, "color": "dark", "y_offset": 55}]},
    ],
)

# --- Stats / Metrics ---
LAYOUTS["stats"] = LayoutTemplate(
    name="数据统计", category="content",
    description="大字 KPI + 进度条/变化百分比，适合数据驱动页",
    elements=[
        {"type": "text", "x": 0, "y": 20, "w": 960, "h": 55, "text": "{title}", "fs": 42, "color": "primary", "bold": True, "align": 2},
        {"type": "line", "x": 300, "y": 78, "w": 360, "h": 3, "color": "secondary"},
        {"type": "stat_cards", "x": 50, "y": 110, "cols": 3, "rows": 1, "gap_x": 30, "item_w": 270, "item_h": 200,
         "fields": ["{value}", "{label}", "{change}"],
         "styles": [{"fs": 48, "color": "primary", "bold": True, "align": 2},
                    {"fs": 16, "color": "gray", "align": 2, "y_offset": 60},
                    {"fs": 14, "color": "{change_color}", "bold": True, "align": 2, "y_offset": 85}]},
        {"type": "sidebar", "x": 50, "y": 340, "w": 860, "h": 150, "color": "light", "text": "{insight}"},
    ],
)

# --- Three Column ---
LAYOUTS["three_col"] = LayoutTemplate(
    name="三栏布局", category="column",
    description="三列垂直排列，适合对比方案/功能列表",
    elements=[
        {"type": "text", "x": 0, "y": 15, "w": 960, "h": 55, "text": "{title}", "fs": 42, "color": "primary", "bold": True, "align": 2},
        {"type": "line", "x": 350, "y": 72, "w": 260, "h": 4, "color": "secondary"},
        {"type": "columns", "x": 35, "y": 100, "cols": 3, "gap": 20, "col_w": 290, "col_h": 420,
         "fields": ["{col_title}", "{col_subtitle}", "{col_content}"],
         "styles": [{"fs": 24, "color": "white", "bold": True, "align": 2, "bg": "accent"},
                    {"fs": 16, "color": "accent_light", "align": 2},
                    {"fs": 15, "color": "light_text", "align": 2, "y_offset": 60}]},
    ],
)

# --- Pipeline / Flow ---
LAYOUTS["pipeline"] = LayoutTemplate(
    name="流程图", category="content",
    description="水平管道流程，带有彩色模块和连接箭头",
    elements=[
        {"type": "text", "x": 0, "y": 15, "w": 960, "h": 55, "text": "{title}", "fs": 42, "color": "primary", "bold": True, "align": 2},
        {"type": "line", "x": 350, "y": 72, "w": 260, "h": 4, "color": "secondary"},
        {"type": "pipeline_items", "x": 35, "y": 100, "count": 6, "gap": 20, "item_w": 140, "item_h": 320,
         "fields": ["{step_num}", "{step_name}", "{step_detail}"],
         "styles": [{"fs": 16, "color": "white", "bold": True, "bg": "{color}", "h": 55},
                    {"fs": 12, "color": "dark", "y_offset": 65, "w": 128, "h": 240}]},
        {"type": "box", "x": 30, "y": 450, "w": 900, "h": 60, "color": "light", "text": "{pipeline_summary}"},
    ],
)

# --- Data Table ---
LAYOUTS["data_table"] = LayoutTemplate(
    name="数据表格", category="content",
    description="左表格 + 右解读，适合排名/对比/指标",
    elements=[
        {"type": "text", "x": 30, "y": 10, "w": 500, "h": 50, "text": "{title}", "fs": 40, "color": "primary", "bold": True},
        {"type": "line", "x": 30, "y": 65, "w": 200, "h": 4, "color": "secondary"},
        {"type": "table", "x": 30, "y": 90, "rows": 7, "cols": 2, "row_h": 42, "col_w": [180, 280],
         "header": ["{col1_name}", "{col2_name}"], "data": "{table_data}"},
        {"type": "box", "x": 530, "y": 90, "w": 400, "h": 410, "color": "light", "text": "{interpretation}"},
    ],
)

# --- Content + Image ---
LAYOUTS["content_image"] = LayoutTemplate(
    name="内容+图片", category="image",
    description="左文右图 / 左图右文，图文并茂",
    elements=[
        {"type": "text", "x": 40, "y": 20, "w": 450, "h": 55, "text": "{title}", "fs": 38, "color": "primary", "bold": True},
        {"type": "line", "x": 40, "y": 80, "w": 120, "h": 3, "color": "secondary"},
        {"type": "text", "x": 40, "y": 100, "w": 440, "h": 400, "text": "{content}", "fs": 18, "color": "dark"},
        {"type": "image_placeholder", "x": 510, "y": 60, "w": 420, "h": 440, "text": "{image_label}"},
    ],
)

# --- Closing ---
LAYOUTS["closing"] = LayoutTemplate(
    name="结语", category="close",
    description="深色全屏 + 总结 + 联系方式/校训",
    elements=[
        {"type": "bg", "color": "primary"},
        {"type": "line", "x": 0, "y": 0, "w": 960, "h": 15, "color": "secondary"},
        {"type": "text", "x": 0, "y": 80, "w": 960, "h": 100, "text": "{summary_title}", "fs": 52, "color": "white", "bold": True, "align": 2},
        {"type": "line", "x": 280, "y": 200, "w": 400, "h": 4, "color": "white"},
        {"type": "text", "x": 60, "y": 230, "w": 840, "h": 200, "text": "{summary_text}", "fs": 18, "color": "light_text", "align": 2},
        {"type": "line", "x": 280, "y": 450, "w": 400, "h": 4, "color": "white"},
        {"type": "text", "x": 0, "y": 470, "w": 960, "h": 50, "text": "{motto}", "fs": 30, "color": "accent_light", "bold": True, "align": 2},
        {"type": "line", "x": 0, "y": 530, "w": 960, "h": 8, "color": "secondary"},
    ],
)


# ============================================================
# Talk-type → recommended slide sequence
# ============================================================

TALK_PRESETS: Dict[str, dict] = {
    "conference": {
        "name": "学术会议",
        "slides": [
            "cover", "toc", "overview", "timeline", "quadrant",
            "grid_cards", "stats", "pipeline", "data_table",
            "content_image", "quadrant", "timeline", "stats", "closing",
        ],
        "rules": {"max_slides": 20, "visual_ratio": 0.65, "key_findings": 2},
    },
    "business": {
        "name": "商务汇报",
        "slides": [
            "cover", "toc", "overview", "stats", "three_col",
            "pipeline", "grid_cards", "data_table", "closing",
        ],
        "rules": {"max_slides": 15, "visual_ratio": 0.5},
    },
    "defense": {
        "name": "论文答辩",
        "slides": [
            "cover", "toc", "overview", "timeline", "quadrant",
            "content_image", "pipeline", "data_table", "stats",
            "quadrant", "timeline", "stats", "closing",
        ],
        "rules": {"max_slides": 65, "visual_ratio": 0.6},
    },
    "school": {
        "name": "学校介绍",
        "slides": [
            "cover", "toc", "overview", "timeline", "three_col",
            "grid_cards", "quadrant", "stats", "closing",
        ],
        "rules": {"max_slides": 14, "visual_ratio": 0.55},
    },
}

# ============================================================
# Registration -- extendable layout system
# ============================================================

def register(layout: LayoutTemplate) -> LayoutTemplate:
    '''Register a new or replacement layout template.'''
    LAYOUTS[layout.name] = layout
    return layout


def unregister(name: str) -> None:
    '''Remove a layout by name.'''
    LAYOUTS.pop(name, None)


def register_talk(name: str, slides: list, rules: dict) -> None:
    '''Register a talk-type preset.'''
    TALK_PRESETS[name] = {"name": name, "slides": slides, "rules": rules}


def resolve_layout_colors(layout: LayoutTemplate, preset) -> LayoutTemplate:
    '''Return a copy of *layout* with colour names resolved to hex strings.

    Args:
        layout: A LayoutTemplate whose elements may reference semantic
                colour names ("primary", "accent", "white", etc.).
        preset: A DesignPreset to resolve against.

    Returns:
        A new LayoutTemplate with all colour names replaced by #RRGGBB hex.
    '''
    from ._colors import resolve_color as _resolve
    import copy

    new_elements = []
    for elem in layout.elements:
        e = copy.deepcopy(elem)
        for key in ("color", "bg", "fg"):
            if key in e and isinstance(e[key], str) and not e[key].startswith("#"):
                try:
                    e[key] = _resolve(e[key], preset)
                except (ValueError, KeyError):
                    pass  # keep original if unresolvable
        new_elements.append(e)
    return LayoutTemplate(
        name=layout.name,
        category=layout.category,
        description=layout.description,
        elements=new_elements,
        structural_rules=layout.structural_rules,
    )



def get_layout(name: str) -> LayoutTemplate:
    '''Return a layout by name.  Raises KeyError if not found.'''
    if name not in LAYOUTS:
        raise KeyError(
            f"Unknown layout '{name}'. Available: {', '.join(LAYOUTS)}"
        )
    return LAYOUTS[name]


def get_talk_preset(name: str) -> dict:
    '''Return a talk-type preset (slide sequence + rules).'''
    if name not in TALK_PRESETS:
        raise KeyError(
            f"Unknown talk type '{name}'. Available: {', '.join(TALK_PRESETS)}"
        )
    return TALK_PRESETS[name]


def list_layouts() -> list:
    '''Return list of available layout names.'''
    return list(LAYOUTS.keys())


def list_talk_types() -> list:
    '''Return list of available talk types.'''
    return list(TALK_PRESETS.keys())
