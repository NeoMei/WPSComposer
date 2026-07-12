"""Agent-friendly COM formatting helpers.

The WPS/Office object model exposes hundreds of dynamic COM properties.  This
module turns the most useful ones into plain JSON-compatible dictionaries and
applies partial patches without resetting properties the caller did not name.
"""

from __future__ import annotations

from ._colors import hex_to_rgb_long, rgb_long_to_hex


def safe_get(obj, name, default=None):
    """Read a COM property without making inspection fail as a whole."""
    try:
        value = getattr(obj, name)
        return default if value is None else value
    except Exception:
        return default


def safe_set(obj, name, value):
    """Set a COM property and return whether the host accepted it."""
    try:
        setattr(obj, name, value)
        return True
    except Exception:
        return False


def as_scalar(value, default=None):
    """Convert common COM scalar values to JSON-safe Python values."""
    if value is None:
        return default
    if isinstance(value, (str, int, float, bool)):
        return value
    try:
        return str(value)
    except Exception:
        return default


def color_hex(value):
    """Return a COM colour long as ``#RRGGBB`` when it is a real colour."""
    try:
        number = int(value)
        if number < 0:  # mixed/automatic sentinel
            return None
        return rgb_long_to_hex(number)
    except Exception:
        return None


def font_snapshot(font):
    if font is None:
        return {}
    color = safe_get(font, "Color")
    if not isinstance(color, (int, float)):
        color = safe_get(color, "RGB")
    return {
        "name": as_scalar(safe_get(font, "Name")),
        "size": as_scalar(safe_get(font, "Size")),
        "bold": as_scalar(safe_get(font, "Bold")),
        "italic": as_scalar(safe_get(font, "Italic")),
        "underline": as_scalar(safe_get(font, "Underline")),
        "strikethrough": as_scalar(
            safe_get(font, "StrikeThrough", safe_get(font, "Strikethrough"))
        ),
        "color": color_hex(color),
    }


def paragraph_snapshot(fmt):
    if fmt is None:
        return {}
    return {
        "alignment": as_scalar(safe_get(fmt, "Alignment")),
        "left_indent": as_scalar(safe_get(fmt, "LeftIndent")),
        "right_indent": as_scalar(safe_get(fmt, "RightIndent")),
        "first_line_indent": as_scalar(safe_get(fmt, "FirstLineIndent")),
        "space_before": as_scalar(safe_get(fmt, "SpaceBefore")),
        "space_after": as_scalar(safe_get(fmt, "SpaceAfter")),
        "line_spacing": as_scalar(safe_get(fmt, "LineSpacing")),
        "line_spacing_rule": as_scalar(safe_get(fmt, "LineSpacingRule")),
        "keep_together": as_scalar(safe_get(fmt, "KeepTogether")),
        "keep_with_next": as_scalar(safe_get(fmt, "KeepWithNext")),
        "page_break_before": as_scalar(safe_get(fmt, "PageBreakBefore")),
        "widow_control": as_scalar(safe_get(fmt, "WidowControl")),
    }


def geometry_snapshot(shape):
    return {
        "left": as_scalar(safe_get(shape, "Left")),
        "top": as_scalar(safe_get(shape, "Top")),
        "width": as_scalar(safe_get(shape, "Width")),
        "height": as_scalar(safe_get(shape, "Height")),
        "rotation": as_scalar(safe_get(shape, "Rotation")),
        "z_order": as_scalar(safe_get(shape, "ZOrderPosition")),
    }


def fill_snapshot(fill):
    if fill is None:
        return {}
    fore = safe_get(fill, "ForeColor")
    back = safe_get(fill, "BackColor")
    return {
        "visible": as_scalar(safe_get(fill, "Visible")),
        "type": as_scalar(safe_get(fill, "Type")),
        "color": color_hex(safe_get(fore, "RGB")),
        "back_color": color_hex(safe_get(back, "RGB")),
        "transparency": as_scalar(safe_get(fill, "Transparency")),
    }


def line_snapshot(line):
    if line is None:
        return {}
    color = safe_get(line, "ForeColor")
    return {
        "visible": as_scalar(safe_get(line, "Visible")),
        "color": color_hex(safe_get(color, "RGB")),
        "weight": as_scalar(safe_get(line, "Weight")),
        "dash_style": as_scalar(safe_get(line, "DashStyle")),
        "transparency": as_scalar(safe_get(line, "Transparency")),
    }


def page_setup_snapshot(page_setup):
    if page_setup is None:
        return {}
    names = (
        "Orientation", "PageWidth", "PageHeight", "TopMargin",
        "BottomMargin", "LeftMargin", "RightMargin", "HeaderDistance",
        "FooterDistance", "Gutter", "PaperSize", "Zoom", "FitToPagesWide",
        "FitToPagesTall", "SlideWidth", "SlideHeight",
    )
    return {
        _snake(name): as_scalar(safe_get(page_setup, name))
        for name in names
        if safe_get(page_setup, name) is not None
    }


def apply_font(font, patch):
    """Apply a partial font patch and return accepted/rejected keys."""
    mapping = {
        "name": "Name", "size": "Size", "bold": "Bold",
        "italic": "Italic", "underline": "Underline",
        "strikethrough": "StrikeThrough",
    }
    accepted, rejected = [], []
    for key, value in patch.items():
        if key == "color":
            color_obj = safe_get(font, "Color")
            if color_obj is not None and not isinstance(color_obj, (int, float)):
                ok = safe_set(color_obj, "RGB", hex_to_rgb_long(value))
            else:
                ok = safe_set(font, "Color", hex_to_rgb_long(value))
        elif key in mapping:
            ok = safe_set(font, mapping[key], value)
        else:
            ok = False
        (accepted if ok else rejected).append(key)
    return {"accepted": accepted, "rejected": rejected}


def apply_paragraph(fmt, patch):
    mapping = {
        "alignment": "Alignment", "left_indent": "LeftIndent",
        "right_indent": "RightIndent", "first_line_indent": "FirstLineIndent",
        "space_before": "SpaceBefore", "space_after": "SpaceAfter",
        "line_spacing": "LineSpacing", "line_spacing_rule": "LineSpacingRule",
        "keep_together": "KeepTogether", "keep_with_next": "KeepWithNext",
        "page_break_before": "PageBreakBefore", "widow_control": "WidowControl",
    }
    return _apply_mapping(fmt, patch, mapping)


def apply_geometry(shape, patch):
    mapping = {
        "left": "Left", "top": "Top", "width": "Width",
        "height": "Height", "rotation": "Rotation",
    }
    return _apply_mapping(shape, patch, mapping)


def apply_fill(fill, patch):
    accepted, rejected = [], []
    for key, value in patch.items():
        ok = False
        if key == "color":
            fore = safe_get(fill, "ForeColor")
            ok = fore is not None and safe_set(fore, "RGB", hex_to_rgb_long(value))
        elif key == "back_color":
            back = safe_get(fill, "BackColor")
            ok = back is not None and safe_set(back, "RGB", hex_to_rgb_long(value))
        elif key == "visible":
            ok = safe_set(fill, "Visible", value)
        elif key == "transparency":
            ok = safe_set(fill, "Transparency", value)
        if ok:
            accepted.append(key)
        else:
            rejected.append(key)
    return {"accepted": accepted, "rejected": rejected}


def apply_line(line, patch):
    accepted, rejected = [], []
    mapping = {"visible": "Visible", "weight": "Weight",
               "dash_style": "DashStyle", "transparency": "Transparency"}
    for key, value in patch.items():
        if key == "color":
            fore = safe_get(line, "ForeColor")
            ok = fore is not None and safe_set(fore, "RGB", hex_to_rgb_long(value))
        elif key in mapping:
            ok = safe_set(line, mapping[key], value)
        else:
            ok = False
        (accepted if ok else rejected).append(key)
    return {"accepted": accepted, "rejected": rejected}


def merge_results(*results):
    return {
        "accepted": [key for result in results for key in result.get("accepted", [])],
        "rejected": [key for result in results for key in result.get("rejected", [])],
    }


def _apply_mapping(obj, patch, mapping, color_keys=None):
    accepted, rejected = [], []
    color_keys = color_keys or {}
    for key, value in patch.items():
        if key in mapping:
            ok = safe_set(obj, mapping[key], value)
        elif key in color_keys:
            ok = safe_set(obj, color_keys[key], hex_to_rgb_long(value))
        else:
            ok = False
        (accepted if ok else rejected).append(key)
    return {"accepted": accepted, "rejected": rejected}


def _snake(name):
    out = []
    for char in name:
        if char.isupper() and out:
            out.append("_")
        out.append(char.lower())
    return "".join(out)
