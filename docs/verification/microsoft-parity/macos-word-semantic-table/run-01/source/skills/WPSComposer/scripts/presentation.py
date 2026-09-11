"""Best-effort desktop presentation of a finalized artifact."""

from __future__ import annotations

from pathlib import Path
import subprocess
import sys
from typing import Union


ArtifactPath = Union[str, Path]


def validate_open_result(open_result: bool) -> None:
    """Reject non-boolean presentation options before callers have effects."""
    if type(open_result) is not bool:
        raise TypeError("open_result must be a bool")


def present_artifact(path: ArtifactPath, *, engine: str | None = None) -> None:
    """Ask the platform default application to open one finalized file."""
    artifact = Path(path).expanduser().resolve()
    if not artifact.is_file():
        raise FileNotFoundError(f"Final artifact is not a file: {artifact}")

    component = {".docx": "writer", ".xlsx": "spreadsheet", ".pptx": "presentation"}.get(artifact.suffix.lower())
    if sys.platform == "darwin":
        if engine in {"wps", "msoffice"} and component is not None:
            application = {"writer": "Microsoft Word", "spreadsheet": "Microsoft Excel", "presentation": "Microsoft PowerPoint"}[component] if engine == "msoffice" else "/Applications/wpsoffice.app"
            argv = ["open", "-a", application, str(artifact)]
        else:
            argv = ["open", str(artifact)]
    elif sys.platform == "win32":
        if engine in {"wps", "msoffice"} and component is not None:
            from .office_engines import engine_executable
            executable = engine_executable(engine, component)
            if not executable:
                raise OSError(f"{engine} executable is unavailable")
            argv = [executable, str(artifact)]
            # Office itself can remain alive for the whole editing session.
            # subprocess.run(timeout=...) would kill it after the handoff.
            subprocess.Popen(
                argv, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL, shell=False,
            )
            return
        else:
            argv = ["explorer.exe", str(artifact)]
    elif sys.platform.startswith("linux"):
        argv = ["xdg-open", str(artifact)]
    else:
        raise OSError(f"Unsupported platform for artifact presentation: {sys.platform}")

    # These platform launchers hand the file to the associated desktop app and
    # then exit.  Wait only for that bounded handoff so a non-zero launcher
    # status remains observable without waiting for the document app to close.
    subprocess.run(
        argv,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        shell=False,
        check=True,
        timeout=10,
    )
