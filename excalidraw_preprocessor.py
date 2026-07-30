#!/usr/bin/env python3
"""
Excalidraw 预处理器：将 .excalidraw.md 文件渲染为 PNG 图片，
并更新 Markdown 文件中的引用。

用法：
    python3 excalidraw_preprocessor.py input.md [--output output.md]
    python3 excalidraw_preprocessor.py input.md --render-only  # 只渲染图片，不修改 md
"""

import argparse
import json
import lzstring
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple, List


def find_excalidraw_refs(md_content: str) -> List[Tuple[str, str, Optional[str], Optional[str]]]:
    """
    找到所有 Excalidraw 引用。
    返回：[(完整匹配，文件路径，宽度，高度), ...]
    """
    # 匹配 ![[path]] 或 ![[path|width]] 或 ![[path|widthxheight]]
    pattern = r'!\[\[([^\]|]+)(?:\|\s*(\d+%?))?(?:\s*x\s*(\d+%?))?\]\]'
    matches = []
    for m in re.finditer(pattern, md_content):
        full_match = m.group(0)
        path = m.group(1).strip()
        width = m.group(2)
        height = m.group(3)
        if path.endswith('.excalidraw.md') or path.endswith('.excalidraw'):
            matches.append((full_match, path, width, height))
    return matches


def extract_compressed_json(excalidraw_path: str) -> Optional[str]:
    """从 .excalidraw.md 文件中提取压缩的 JSON 数据。"""
    with open(excalidraw_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 匹配 ```compressed-json ... ``` 块
    match = re.search(r'```compressed-json\s*\n(.*?)\n```', content, re.DOTALL)
    if not match:
        return None
    
    compressed_data = match.group(1).strip()
    # 移除空白字符（LZString 不处理换行）
    compressed_data = re.sub(r'\s+', '', compressed_data)
    return compressed_data


def decompress_excalidraw(compressed_data: str) -> Optional[dict]:
    """解压 Excalidraw 数据。"""
    try:
        lz = lzstring.LZString()
        decompressed = lz.decompressFromBase64(compressed_data)
        if not decompressed:
            return None
        return json.loads(decompressed)
    except Exception as e:
        print(f"解压失败：{e}", file=sys.stderr)
        return None


def render_excalidraw_to_png(excalidraw_path: str, output_path: str, width: Optional[int] = None) -> bool:
    """
    使用 Excalidraw 将文件渲染为 PNG。
    
    方案：使用 excalidraw-cli 或 npx @excalidraw/excalidraw-cli
    如果不可用，回退到使用浏览器渲染。
    """
    # 方案 1: 使用 excalidraw-cli（如果已安装）
    try:
        cmd = ['npx', '@excalidraw/excalidraw-cli', 'export', 
               '--input', excalidraw_path,
               '--output', output_path,
               '--type', 'png']
        if width:
            cmd.extend(['--width', str(width)])
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode == 0:
            return True
        print(f"excalidraw-cli 失败：{result.stderr}", file=sys.stderr)
    except FileNotFoundError:
        print("excalidraw-cli 未安装", file=sys.stderr)
    except subprocess.TimeoutExpired:
        print("excalidraw-cli 超时", file=sys.stderr)
    except Exception as e:
        print(f"excalidraw-cli 错误：{e}", file=sys.stderr)
    
    # 方案 2: 使用 Python 脚本 + Playwright 渲染
    return render_with_playwright(excalidraw_path, output_path, width)


def render_with_playwright(excalidraw_path: str, output_path: str, width: Optional[int] = None, max_retries: int = 3) -> bool:
    """使用 Playwright + Excalidraw 渲染为 PNG。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright 未安装，请运行：pip install playwright && playwright install chromium", file=sys.stderr)
        return False
    
    # 读取并解压 Excalidraw 数据
    compressed_data = extract_compressed_json(excalidraw_path)
    if not compressed_data:
        print(f"无法从 {excalidraw_path} 提取数据", file=sys.stderr)
        return False
    
    scene_data = decompress_excalidraw(compressed_data)
    if not scene_data:
        print(f"无法解压 {excalidraw_path}", file=sys.stderr)
        return False
    
    # 计算渲染尺寸
    elements = scene_data.get('elements', [])
    if not elements:
        return False
    
    min_x = min_y = float('inf')
    max_x = max_y = float('-inf')
    
    for el in elements:
        if el.get('isDeleted'):
            continue
        x, y = el.get('x', 0), el.get('y', 0)
        w, h = el.get('width', 0), el.get('height', 0)
        min_x = min(min_x, x)
        min_y = min(min_y, y)
        max_x = max(max_x, x + w)
        max_y = max(max_y, y + h)
    
    if min_x == float('inf'):
        return False
    
    padding = 40
    render_width = int(max_x - min_x + padding * 2)
    render_height = int(max_y - min_y + padding * 2)
    
    if width:
        render_width = width
        render_height = int(render_height * (width / (max_x - min_x + padding * 2)))
    
    # 重试渲染
    for attempt in range(max_retries):
        success = _try_render_once(scene_data, render_width, render_height, output_path)
        if success:
            return True
        print(f"  第 {attempt + 1} 次渲染失败，重试...", file=sys.stderr)
    
    return False


def _try_render_once(scene_data: dict, render_width: int, render_height: int, output_path: str) -> bool:
    """尝试渲染一次 Excalidraw 场景。"""
    from playwright.sync_api import sync_playwright
    
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': render_width + 200, 'height': render_height + 200})
        
        # 创建包含 Excalidraw 的 HTML - 使用动态加载和导出 API
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
                // 动态加载脚本
                function loadScript(src) {{
                    return new Promise((resolve, reject) => {{
                        const script = document.createElement('script');
                        script.src = src;
                        script.onload = resolve;
                        script.onerror = reject;
                        document.head.appendChild(script);
                    }});
                }}
                
                const sceneData = {json.dumps(scene_data)};
                
                async function exportScene() {{
                    try {{
                        // 加载依赖
                        await loadScript('https://unpkg.com/react@18/umd/react.production.min.js');
                        await loadScript('https://unpkg.com/react-dom@18/umd/react-dom.production.min.js');
                        await loadScript('https://unpkg.com/@excalidraw/excalidraw@0.17.6/dist/excalidraw.production.min.js');
                        
                        // 检查 API
                        if (!window.ExcalidrawLib || !window.ExcalidrawLib.exportToCanvas) {{
                            console.error('ExcalidrawLib.exportToCanvas not available');
                            window.__exportComplete = false;
                            return;
                        }}
                        
                        // 导出为 canvas
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
                        
                        // 显示 canvas
                        const container = document.getElementById('canvas-container');
                        container.innerHTML = '';
                        container.appendChild(canvas);
                        
                        window.__exportComplete = true;
                        console.log('Export complete');
                    }} catch (e) {{
                        console.error('Export failed:', e);
                        window.__exportComplete = false;
                    }}
                }}
                
                // 执行导出
                exportScene();
            </script>
        </body>
        </html>
        """
        
        page.set_content(html_content)
        
        # 等待 Excalidraw 加载和导出完成（需要更长时间）
        page.wait_for_timeout(15000)
        
        # 检查是否导出完成
        export_complete = page.evaluate('() => window.__exportComplete || false')
        
        if not export_complete:
            print(f"  导出未完成", file=sys.stderr)
            browser.close()
            return False
        
        # 等待 canvas 渲染完成
        page.wait_for_timeout(1000)
        
        # 只截图 canvas 容器区域，避免工具栏
        element = page.query_selector('#canvas-container')
        if element:
            element.screenshot(path=output_path)
        else:
            page.screenshot(path=output_path, full_page=True)
        
        browser.close()
        
        # 检查生成的图片是否为空白
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            if file_size < 10000:  # 小于 10KB 可能是空白
                print(f"  图片太小 ({file_size} bytes)，可能是空白", file=sys.stderr)
                return False
            return True
        
        return False


def process_markdown_file(md_path: str, output_path: Optional[str] = None, render_only: bool = False) -> str:
    """
    处理 Markdown 文件，将 Excalidraw 引用替换为图片引用。
    
    返回：更新后的 Markdown 内容
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    md_dir = Path(md_path).parent
    refs = find_excalidraw_refs(content)
    
    if not refs:
        print(f"未找到 Excalidraw 引用")
        return content
    
    print(f"找到 {len(refs)} 个 Excalidraw 引用")
    
    for full_match, excalidraw_path, width_str, height_str in refs:
        # 解析尺寸
        width = None
        if width_str and '%' not in width_str:
            width = int(width_str)
        
        # 解析 Excalidraw 文件路径（相对于 md 文件或 vault 根目录）
        if not os.path.isabs(excalidraw_path):
            # 先尝试相对于 md 文件
            abs_path = md_dir / excalidraw_path
            if not abs_path.exists():
                # 尝试添加 .md 扩展名
                abs_path = md_dir / f"{excalidraw_path}.md"
            if not abs_path.exists():
                # 尝试从 vault 根目录（向上查找 .obsidian）
                vault_root = find_vault_root(md_dir)
                if vault_root:
                    abs_path = vault_root / excalidraw_path
                    if not abs_path.exists():
                        abs_path = vault_root / f"{excalidraw_path}.md"
        else:
            abs_path = Path(excalidraw_path)
        
        if not abs_path.exists():
            print(f"警告：找不到文件 {excalidraw_path}", file=sys.stderr)
            continue
        
        # 生成 PNG 文件名
        png_name = abs_path.stem + '.png'
        png_path = md_dir / png_name
        
        print(f"渲染：{abs_path.name} -> {png_name}")
        
        # 渲染为 PNG
        success = render_excalidraw_to_png(str(abs_path), str(png_path), width)
        
        if success:
            # 替换引用
            if render_only:
                print(f"  ✅ 已渲染为 {png_path}")
            else:
                # 替换为 Markdown 图片语法（使用绝对路径，避免 WpsComposer 找不到文件）
                new_ref = f"![{abs_path.stem}]({png_path.absolute()})"
                content = content.replace(full_match, new_ref)
                print(f"  ✅ 已替换为 {new_ref}")
        else:
            print(f"  ❌ 渲染失败", file=sys.stderr)
    
    if not render_only:
        # 写入更新后的文件
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"已保存到：{output_path}")
        else:
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"已更新：{md_path}")
    
    return content


def find_vault_root(start_dir: Path) -> Optional[Path]:
    """查找 Obsidian vault 根目录。"""
    current = start_dir.absolute()
    while True:
        if (current / '.obsidian').is_dir():
            return current
        parent = current.parent
        if parent == current:
            return None
        current = parent


def main():
    parser = argparse.ArgumentParser(description='Excalidraw 预处理器')
    parser.add_argument('input', help='输入的 Markdown 文件')
    parser.add_argument('--output', '-o', help='输出文件路径（默认覆盖输入文件）')
    parser.add_argument('--render-only', action='store_true', help='只渲染图片，不修改 Markdown')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"错误：文件不存在 {args.input}", file=sys.stderr)
        sys.exit(1)
    
    process_markdown_file(args.input, args.output, args.render_only)


if __name__ == '__main__':
    main()
