"""Excalidraw plugin for WPSComposer.

Renders .excalidraw.md files to PNG images and replaces
Obsidian wikilink references with standard Markdown image syntax.

Usage::

    from skills.WPSComposer import generate
    generate("input.md", format="pdf", plugins=["excalidraw"])
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Optional


# Match Obsidian wikilink image syntax: ![[path]] or ![[path|width]] or ![[path|widthxheight]]
_WIKILINK_IMAGE_RE = re.compile(
    r"!\[\[([^\]|]+)(?:\|\s*(\d+%?))?(?:\s*x\s*(\d+%?))?\]\]"
)


def excalidraw_plugin(content: str, base_dir: str) -> str:
    """Plugin entry point: render Excalidraw references to PNG images.

    Args:
        content: Raw Markdown text.
        base_dir: Base directory for resolving relative paths.

    Returns:
        Modified Markdown text with Excalidraw references replaced
        by standard Markdown image syntax pointing to PNG files.
    """
    refs = _find_excalidraw_refs(content)
    if not refs:
        return content

    print(f"[excalidraw] Found {len(refs)} Excalidraw reference(s)", file=sys.stderr)

    for full_match, excalidraw_rel_path, width_str, height_str in refs:
        width = _parse_dimension(width_str)

        # Resolve the Excalidraw file path
        abs_path = _resolve_excalidraw_path(excalidraw_rel_path, base_dir)
        if abs_path is None:
            print(
                f"[excalidraw] Warning: cannot find {excalidraw_rel_path}",
                file=sys.stderr,
            )
            continue

        # Determine output PNG path (same directory as the Excalidraw file)
        png_path = abs_path.parent / f"{abs_path.stem}.png"

        print(
            f"[excalidraw] Rendering {abs_path.name} -> {png_path.name}",
            file=sys.stderr,
        )

        success = _render_excalidraw_to_png(str(abs_path), str(png_path), width)

        if success:
            # Replace wikilink with standard Markdown image syntax (absolute path)
            new_ref = f"![{abs_path.stem}]({png_path})"
            content = content.replace(full_match, new_ref)
            print(f"[excalidraw]   OK -> {new_ref}", file=sys.stderr)
        else:
            print(
                f"[excalidraw]   FAILED to render {abs_path.name}",
                file=sys.stderr,
            )

    return content


def _find_excalidraw_refs(content: str):
    """Find all Excalidraw wikilink references in Markdown content.

    Returns list of (full_match, path, width_str, height_str).
    """
    refs = []
    for m in _WIKILINK_IMAGE_RE.finditer(content):
        path = m.group(1).strip()
        if path.endswith(".excalidraw.md") or path.endswith(".excalidraw"):
            refs.append((m.group(0), path, m.group(2), m.group(3)))
    return refs


def _parse_dimension(s: Optional[str]) -> Optional[int]:
    """Parse a dimension string like '300' or '100%'. Returns int or None."""
    if not s:
        return None
    s = s.strip()
    if "%" in s:
        return None  # percentage not supported for PNG rendering
    try:
        return int(s)
    except ValueError:
        return None


def _resolve_excalidraw_path(rel_path: str, base_dir: str) -> Optional[Path]:
    """Resolve an Excalidraw file path relative to base_dir or vault root."""
    base = Path(base_dir)

    # Try relative to base_dir
    candidate = base / rel_path
    if candidate.exists():
        return candidate

    # Try with .md extension
    candidate_md = base / f"{rel_path}.md"
    if candidate_md.exists():
        return candidate_md

    # Try vault root (walk up looking for .obsidian)
    vault_root = _find_vault_root(base)
    if vault_root:
        candidate = vault_root / rel_path
        if candidate.exists():
            return candidate
        candidate_md = vault_root / f"{rel_path}.md"
        if candidate_md.exists():
            return candidate_md

    return None


def _find_vault_root(start: Path) -> Optional[Path]:
    """Find Obsidian vault root by looking for .obsidian directory."""
    current = start.absolute()
    while True:
        if (current / ".obsidian").is_dir():
            return current
        parent = current.parent
        if parent == current:
            return None
        current = parent


def _render_excalidraw_to_png(
    excalidraw_path: str,
    output_path: str,
    width: Optional[int] = None,
    max_retries: int = 3,
) -> bool:
    """Render an Excalidraw file to PNG using Playwright.

    Uses Excalidraw's exportToCanvas API to produce a clean image
    without the editor UI (toolbar, sidebar, etc.).
    """
    compressed_data = _extract_compressed_json(excalidraw_path)
    if not compressed_data:
        print(f"  No compressed-json block found", file=sys.stderr)
        return False

    scene_data = _decompress_excalidraw(compressed_data)
    if not scene_data:
        print(f"  Failed to decompress", file=sys.stderr)
        return False

    elements = scene_data.get("elements", [])
    if not elements:
        return False

    # Calculate bounding box
    min_x = min_y = float("inf")
    max_x = max_y = float("-inf")
    for el in elements:
        if el.get("isDeleted"):
            continue
        x, y = el.get("x", 0), el.get("y", 0)
        w, h = el.get("width", 0), el.get("height", 0)
        min_x = min(min_x, x)
        min_y = min(min_y, y)
        max_x = max(max_x, x + w)
        max_y = max(max_y, y + h)

    if min_x == float("inf"):
        return False

    padding = 40
    original_width = int(max_x - min_x + padding * 2)
    original_height = int(max_y - min_y + padding * 2)
    render_width = original_width
    render_height = original_height

    if width and original_width > 0:
        render_width = width
        render_height = int(original_height * (width / original_width))
    elif width:
        # Fallback: keep original dimensions if original_width is 0
        pass

    for attempt in range(max_retries):
        if _try_render_once(scene_data, render_width, render_height, output_path):
            return True
        print(f"  Attempt {attempt + 1} failed, retrying...", file=sys.stderr)

    return False


def _extract_compressed_json(excalidraw_path: str) -> Optional[str]:
    """Extract compressed JSON from a .excalidraw.md file."""
    with open(excalidraw_path, "r", encoding="utf-8") as f:
        content = f.read()
    match = re.search(r"```compressed-json\s*\n(.*?)\n```", content, re.DOTALL)
    if not match:
        return None
    compressed = match.group(1).strip()
    return re.sub(r"\s+", "", compressed)


def _decompress_excalidraw(compressed_data: str) -> Optional[dict]:
    """Decompress Excalidraw LZString data."""
    try:
        import lzstring
    except ImportError:
        print(
            "  lzstring not installed. Run: pip install lzstring",
            file=sys.stderr,
        )
        return None
    try:
        lz = lzstring.LZString()
        decompressed = lz.decompressFromBase64(compressed_data)
        if not decompressed:
            return None
        return json.loads(decompressed)
    except Exception as e:
        print(f"  Decompression error: {e}", file=sys.stderr)
        return None


def _try_render_once(
    scene_data: dict,
    render_width: int,
    render_height: int,
    output_path: str,
) -> bool:
    """Attempt one Excalidraw render via Playwright."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print(
            "  Playwright not installed. Run: pip install playwright && playwright install chromium",
            file=sys.stderr,
        )
        return False

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(
            viewport={"width": render_width + 200, "height": render_height + 200}
        )

        # Escape scene_data for safe embedding in HTML/JS
        scene_data_json = json.dumps(scene_data)
        # Escape </script> to prevent breaking out of the script tag
        scene_data_json = scene_data_json.replace("</script>", "<\\/script>")
        scene_data_json = scene_data_json.replace("</", "<\\/")

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ margin: 0; padding: 0; background: white; }}
                #canvas-container {{ width: {render_width}px; height: {render_height}px; }}
            </style>
        </head>
        <body>
            <div id="canvas-container"></div>
            <script>
                function loadScript(src) {{
                    return new Promise((resolve, reject) => {{
                        const script = document.createElement('script');
                        script.src = src;
                        script.onload = resolve;
                        script.onerror = reject;
                        document.head.appendChild(script);
                    }});
                }}

                const sceneData = {scene_data_json};

                async function exportScene() {{
                    try {{
                        await loadScript('https://unpkg.com/react@18/umd/react.production.min.js');
                        await loadScript('https://unpkg.com/react-dom@18/umd/react-dom.production.min.js');
                        await loadScript('https://unpkg.com/@excalidraw/excalidraw@0.17.6/dist/excalidraw.production.min.js');

                        if (!window.ExcalidrawLib || !window.ExcalidrawLib.exportToCanvas) {{
                            window.__exportComplete = false;
                            return;
                        }}

                        const canvas = await window.ExcalidrawLib.exportToCanvas({{
                            elements: sceneData.elements || [],
                            appState: {{
                                viewBackgroundColor: '#ffffff',
                                exportWithDarkMode: false,
                            }},
                            files: sceneData.files || {{}},
                            scale: 1,
                            renderEmbeddings: true,
                        }});

                        const container = document.getElementById('canvas-container');
                        container.innerHTML = '';
                        container.appendChild(canvas);
                        window.__exportComplete = true;
                    }} catch (e) {{
                        console.error('Export failed:', e);
                        window.__exportComplete = false;
                    }}
                }}

                exportScene();
            </script>
        </body>
        </html>
        """

        page.set_content(html_content, timeout=60000)  # 60 second timeout for content loading
        page.wait_for_timeout(15000)

        export_complete = page.evaluate("() => window.__exportComplete || false")
        if not export_complete:
            browser.close()
            return False

        page.wait_for_timeout(1000)

        element = page.query_selector("#canvas-container")
        if element:
            element.screenshot(path=output_path)
        else:
            page.screenshot(path=output_path, full_page=True)

        browser.close()

        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            if file_size < 10000:
                print(f"  Image too small ({file_size} bytes), likely blank", file=sys.stderr)
                return False
            return True

        return False
