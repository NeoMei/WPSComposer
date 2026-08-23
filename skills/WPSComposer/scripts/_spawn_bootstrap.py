"""No-op main module used by spawned validator workers in interactive hosts.

Python's spawn preparation insists on reloading ``__main__.__file__``.  REPL,
stdin, and some notebook hosts expose a synthetic name instead of a file, so
artifact validation temporarily points preparation at this import-safe module.
"""

from __future__ import annotations

