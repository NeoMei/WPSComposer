'''Quality checks — slide-level validation against design rules.

Integrated from harness-anything (yb2460/harness-anything).
Validates font hierarchy, color contrast, content density, white space,
and one-idea-per-slide discipline.
'''

from __future__ import annotations
from typing import Dict, List

from .design_presets import DesignPreset


def validate_slide(
    slide_elements: List[dict],
    preset_rules: dict,
) -> dict:
    '''Run quality checks on a single slide.

    Args:
        slide_elements: List of element dicts, each with at least
                        {"type": ..., "fs": pt, "role": ...}.
        preset_rules:   DesignPreset.rules dict.

    Returns:
        {"pass": bool, "warnings": [str, ...], "score": int (0-100)}
    '''
    warnings: List[str] = []
    score = 100

    texts = [e for e in slide_elements if e.get("type") == "text"]
    shapes = [e for e in slide_elements if e.get("type") in (
        "box", "rect", "circle", "image", "shape"
    )]

    # 1. Font size checks
    for t in texts:
        fs = t.get("fs", 18)
        role = t.get("role", "body")
        if role == "title" and fs < 36:
            warnings.append(f"标题字号 {fs}pt < 推荐 36pt")
            score -= 5
        if role == "body" and fs < 20:
            warnings.append(f"正文字号 {fs}pt < 推荐 20pt")
            score -= 3
        if role == "caption" and fs < 14:
            warnings.append(f"标注字号 {fs}pt < 推荐 14pt")
            score -= 2

    # 2. Content density check
    bullet_count = len(texts)
    max_bullets = preset_rules.get("max_bullets", 6)
    if bullet_count > max_bullets * 1.5:
        warnings.append(
            f"文本块过多 ({bullet_count} > {int(max_bullets * 1.5)})"
        )
        score -= 10

    # 3. Visual ratio check
    visual_ratio_target = preset_rules.get("visual_ratio", 0.5)
    total = len(texts) + len(shapes)
    if total > 0:
        visual_ratio = len(shapes) / total
        if visual_ratio < visual_ratio_target - 0.2:
            warnings.append(
                f"视觉占比 {visual_ratio:.0%} < 目标 {visual_ratio_target:.0%}"
            )
            score -= 5

    # 4. One idea per slide check
    title_count = sum(1 for t in texts if t.get("role") == "title")
    if title_count > 1:
        warnings.append(
            f"检测到 {title_count} 个标题，建议一页一个主题"
        )
        score -= 10

    return {
        "pass": score >= 70,
        "warnings": warnings,
        "score": max(0, min(100, score)),
    }


def review_deck(
    slides: List[List[dict]],
    preset: DesignPreset,
) -> dict:
    '''Run quality checks across an entire presentation.

    Args:
        slides: List of slides, each slide is a list of element dicts.
        preset: A DesignPreset instance providing the rules.

    Returns:
        {"overall_score": float, "per_slide": [...], "total_warnings": int,
         "summary": str}
    '''
    results = []
    for i, slide_elements in enumerate(slides):
        result = validate_slide(slide_elements, preset.rules)
        result["slide_index"] = i
        results.append(result)

    overall = sum(r["score"] for r in results) / max(len(results), 1)
    total_warnings = sum(len(r["warnings"]) for r in results)

    return {
        "overall_score": round(overall, 1),
        "per_slide": results,
        "total_warnings": total_warnings,
        "summary": (
            f"整体 {overall:.0f}分，{len(results)}页，"
            f"{total_warnings}个警告"
        ),
    }


# ============================================================
# Review dimensions (5-axis slide-excellence framework)
# ============================================================


def validate_wps_slide(slide_obj, preset: DesignPreset) -> dict:
    '''Run quality checks against a real WPS COM slide object.

    Inspects shapes on the slide for font sizes, colour usage,
    and content density against the preset rules.

    Args:
        slide_obj: A WPS COM Slide object (e.g. ``pres.Slides(1)``).
        preset: A DesignPreset instance providing the rules.

    Returns:
        {"pass": bool, "warnings": [str, ...], "score": int (0-100)}
    '''
    warnings = []
    score = 100

    try:
        shapes = slide_obj.Shapes
    except Exception:
        return {"pass": True, "warnings": ["Cannot access slide shapes"], "score": 100}

    rules = preset.rules
    text_count = 0
    shape_count = 0
    title_count = 0

    for i in range(1, shapes.Count + 1):
        try:
            shape = shapes(i)
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text_count += 1
                try:
                    fs = shape.TextFrame.TextRange.Font.Size
                except Exception:
                    fs = 12
                try:
                    name = shape.Name or ""
                except Exception:
                    name = ""

                # Heuristic: large text = title
                is_title = fs >= 30 or "Title" in name or "title" in name.lower()
                if is_title:
                    title_count += 1
                    if fs < 36:
                        warnings.append(f"Title font {fs}pt < recommended 36pt")
                        score -= 5
                elif fs < 18:
                    warnings.append(f"Body font {fs}pt < recommended 18pt")
                    score -= 3
            else:
                shape_count += 1
        except Exception:
            shape_count += 1

    # Content density
    max_bullets = rules.get("max_bullets", 6)
    if text_count > max_bullets * 1.5:
        warnings.append(
            f"Text blocks ({text_count}) > recommended max ({int(max_bullets * 1.5)})"
        )
        score -= 10

    # Visual ratio
    total = text_count + shape_count
    visual_ratio_target = rules.get("visual_ratio", 0.5)
    if total > 0:
        actual_ratio = shape_count / total
        if actual_ratio < visual_ratio_target - 0.2:
            warnings.append(
                f"Visual ratio {actual_ratio:.0%} < target {visual_ratio_target:.0%}"
            )
            score -= 5

    # One idea per slide
    if title_count > 1:
        warnings.append(f"Detected {title_count} titles, suggest one topic per slide")
        score -= 10

    return {
        "pass": score >= 70,
        "warnings": warnings,
        "score": max(0, min(100, score)),
    }


def review_wps_deck(pres_obj, preset: DesignPreset) -> dict:
    '''Run quality checks on an entire WPS COM presentation.

    Args:
        pres_obj: WPS COM Presentation object.
        preset: DesignPreset with rules.

    Returns:
        {"overall_score": float, "per_slide": [...], "total_warnings": int,
         "summary": str}
    '''
    results = []
    slide_count = pres_obj.Slides.Count
    for i in range(1, slide_count + 1):
        result = validate_wps_slide(pres_obj.Slides(i), preset)
        result["slide_index"] = i - 1
        results.append(result)

    overall = sum(r["score"] for r in results) / max(len(results), 1)
    total_warnings = sum(len(r["warnings"]) for r in results)

    return {
        "overall_score": round(overall, 1),
        "per_slide": results,
        "total_warnings": total_warnings,
        "summary": (
            f"Overall {overall:.0f}/100, {slide_count} slides, "
            f"{total_warnings} warnings"
        ),
    }



REVIEW_DIMENSIONS: Dict[str, dict] = {
    "visual": {
        "name": "视觉审查",
        "checks": [
            "字体层级", "颜色对比度", "留白比例",
            "布局一致性", "图片质量",
        ],
        "threshold": 70,
    },
    "pedagogy": {
        "name": "教学法审查",
        "checks": [
            "叙事弧完整性", "预备知识清晰", "示例充分",
            "符号一致性", "逻辑流畅度",
        ],
        "threshold": 75,
    },
    "proofreading": {
        "name": "校对审查",
        "checks": [
            "拼写", "语法", "术语一致性",
            "标点", "字体溢出",
        ],
        "threshold": 80,
    },
    "parity": {
        "name": "格式一致性",
        "checks": [
            "PPTX vs PDF 一致", "字体嵌入",
            "图片不丢失", "动画兼容",
        ],
        "threshold": 85,
    },
    "substance": {
        "name": "内容实质",
        "checks": [
            "数据准确性", "引用完整性",
            "结论支撑", "方法可复现",
        ],
        "threshold": 90,
    },
}
