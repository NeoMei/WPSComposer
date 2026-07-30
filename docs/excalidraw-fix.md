---
title: WpsComposer Excalidraw 嵌入修复
created: 2026-07-30
status: 草案
tags:
  - WpsComposer
  - bug
  - Excalidraw
  - Obsidian
---

# WpsComposer Excalidraw 嵌入修复

## 问题描述

Markdown 文档中内嵌的 Excalidraw 图形（`![[...excalidraw]]`）在 WpsComposer 转换时被输出为文件名字符串，而不是渲染为图片。

**正确展现形式**：图形的图片嵌入文档。

## 根因分析

### 1. Obsidian wikilink 图片语法未被解析

当前 `_IMAGE_RE` 只匹配标准 Markdown 图片语法：
```python
_IMAGE_RE = re.compile(
    r"!\[([^\]]*)\]\(\s*(?:<([^>]+)>|([^\s)]+))"
    r"(?:\s+(?:\"[^\"]*\"|'[^']*'))?\s*\)"
)
```

这匹配 `![alt](url)` 格式，但 **不匹配** Obsidian 的 `![[path]]` wikilink 格式。

### 2. Excalidraw 文件需要特殊处理

Excalidraw 文件是 `.excalidraw.md` 格式（Markdown 内嵌 JSON），不是直接的图片文件。需要：
1. 解析 `.excalidraw.md` 文件，提取其中的 JSON 数据
2. 将 JSON 渲染为 SVG/PNG 图片
3. 将图片嵌入文档

### 3. 当前行为

当遇到 `![[xxx.excalidraw]]` 时：
- `_IMAGE_RE` 不匹配
- 被当作普通段落文本处理
- 输出为字符串 `![[xxx.excalidraw]]`

## 修复方案

### 方案 A：支持 Obsidian wikilink 图片（基础）

添加对 `![[path]]` 语法的解析，支持直接嵌入图片文件（PNG/JPG/SVG 等）。

```python
# 新增 wikilink 图片正则
_WIKILINK_IMAGE_RE = re.compile(r"^!\[\[([^\]]+)\]\](?:\|(\d+))?(?:\|(\d+))?$")

def _is_wikilink_image(line: str) -> Optional[re.Match]:
    """Check if line is an Obsidian wikilink image reference."""
    return _WIKILINK_IMAGE_RE.match(line.strip())
```

在块级解析中添加处理：
```python
# 在 parse() 函数的块级解析循环中添加
wikilink_match = _is_wikilink_image(line)
if wikilink_match:
    path = wikilink_match.group(1).strip()
    width = int(wikilink_match.group(2)) if wikilink_match.group(2) else None
    height = int(wikilink_match.group(3)) if wikilink_match.group(3) else None
    current_section = _ensure_section(current_section, sections)
    current_section.elements.append(ImageBlock(
        path=_resolve_image_path(path, base_dir),
        alt=os.path.basename(path),
        width=width,
        height=height,
    ))
    i += 1
    continue
```

### 方案 B：Excalidraw 文件渲染（完整）

对于 `.excalidraw.md` 文件，需要额外处理：

1. **检测 Excalidraw 文件**：
```python
def _is_excalidraw_file(path: str) -> bool:
    return path.endswith('.excalidraw.md') or path.endswith('.excalidraw')
```

2. **解析 Excalidraw JSON**：
```python
import json

def _parse_excalidraw_json(md_path: str) -> Optional[dict]:
    """Extract Excalidraw JSON from .excalidraw.md file."""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Excalidraw.md files have JSON between ```json blocks or as raw JSON
    # after YAML frontmatter
    json_match = re.search(r'```json\s*\n(.*?)\n```', content, re.DOTALL)
    if json_match:
        return json.loads(json_match.group(1))
    
    # Try to find JSON object directly
    json_match = re.search(r'\{.*"type"\s*:\s*"excalidraw".*\}', content, re.DOTALL)
    if json_match:
        return json.loads(json_match.group(0))
    
    return None
```

3. **渲染为 SVG**：
使用 Excalidraw 的官方渲染库或简化版渲染器将 JSON 转换为 SVG。

4. **嵌入 SVG 到文档**：
```python
def _render_excalidraw(w, json_data, width=None, height=None):
    """Render Excalidraw JSON as embedded SVG."""
    svg_content = render_excalidraw_to_svg(json_data)
    w.add_svg_block(svg_content, width=width, height=height)
```

### 推荐实施顺序

1. **先实施方案 A**：支持 `![[path.png]]` 等直接图片引用
2. **再实施方案 B**：支持 Excalidraw 文件渲染

## 测试用例

### 测试 1：标准 Markdown 图片（现有功能）
```markdown
![alt text](image.png)
```
✅ 应正常渲染

### 测试 2：Obsidian wikilink 图片（待修复）
```markdown
![[image.png]]
![[image.png|300]]
![[image.png|300|200]]
```
 当前输出字符串，应渲染为图片

### 测试 3：Excalidraw 文件（待修复）
```markdown
![[diagram.excalidraw]]
![[diagram.excalidraw|500]]
```
 当前输出字符串，应渲染为 SVG 图片

### 测试 4：Excalidraw 带尺寸
```markdown
![[diagram.excalidraw|600|400]]
```
❌ 当前输出字符串，应渲染为 600x400 的 SVG 图片

## 相关文件

- `/Users/neomei/项目/WpsComposer/skills/WPSComposer/scripts/md_parser.py`
- `/Users/neomei/项目/WpsComposer/skills/WPSComposer/scripts/document_model.py`
- `/Users/neomei/项目/WpsComposer/skills/WPSComposer/scripts/renderers/writer_renderer.py`

## 参考

- Obsidian wikilink 语法：`![[filename]]`、`![[filename|alias]]`、`![[filename|width]]`、`![[filename|widthxheight]]`
- Excalidraw 文件格式：`.excalidraw.md`（Markdown + JSON）或 `.excalidraw`（纯 JSON）
