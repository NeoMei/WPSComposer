"""WPSComposer plugin system.

Plugins are callables that preprocess Markdown content before parsing.
Each plugin receives the raw Markdown text and base directory,
and returns the modified Markdown text.

Plugin interface::

    def my_plugin(content: str, base_dir: str) -> str:
        # modify content
        return content

Built-in plugins:
    - excalidraw: Render .excalidraw.md files to PNG images
"""

from __future__ import annotations

from typing import Callable, Dict, List, Optional

# Plugin type: takes (content, base_dir) -> modified content
PluginFunc = Callable[[str, str], str]

# Registry of built-in plugins
_BUILTIN_PLUGINS: Dict[str, PluginFunc] = {}

# Lazy-loadable built-in plugins (name -> module_path)
_LAZY_PLUGINS: Dict[str, str] = {
    "excalidraw": ".excalidraw",
}


def register_plugin(name: str, func: PluginFunc) -> None:
    """Register a plugin by name."""
    _BUILTIN_PLUGINS[name] = func


def get_plugin(name: str) -> Optional[PluginFunc]:
    """Get a plugin by name. Returns None if not found."""
    plugin = _BUILTIN_PLUGINS.get(name)
    if plugin is None and name in _LAZY_PLUGINS:
        # Try to lazy-load
        module_path = _LAZY_PLUGINS[name]
        try:
            import importlib
            module = importlib.import_module(module_path, package=__package__)
            plugin_func = getattr(module, f"{name}_plugin")
            register_plugin(name, plugin_func)
            return plugin_func
        except (ImportError, AttributeError) as exc:
            raise RuntimeError(f"Failed to load plugin '{name}'") from exc
    return plugin


def list_plugins() -> List[str]:
    """Return names of all available plugins (including lazy-loadable)."""
    return list(set(list(_BUILTIN_PLUGINS.keys()) + list(_LAZY_PLUGINS.keys())))


def run_plugins(
    content: str,
    base_dir: str,
    plugin_names: Optional[List[str]] = None,
) -> str:
    """Run a list of plugins on Markdown content.

    Args:
        content: Raw Markdown text.
        base_dir: Base directory for resolving relative paths.
        plugin_names: List of plugin names to run. If None, no plugins run.

    Returns:
        Modified Markdown text after all plugins have run.
    """
    if not plugin_names:
        return content

    for name in plugin_names:
        plugin = get_plugin(name)
        if plugin is None:
            raise ValueError(
                f"Unknown plugin '{name}'. Available plugins: {list_plugins()}"
            )
        content = plugin(content, base_dir)
        if not isinstance(content, str):
            raise TypeError(f"Plugin '{name}' must return Markdown text")

    return content
