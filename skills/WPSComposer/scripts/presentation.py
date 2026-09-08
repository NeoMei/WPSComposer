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

    if sys.platform == "darwin":
        argv = ["open", "-a", "Microsoft Word", str(artifact)] if engine == "msoffice" and artifact.suffix.lower() == ".docx" else ["open", str(artifact)]
    elif sys.platform == "win32":
        if engine == "msoffice" and artifact.suffix.lower() == ".docx":
            from .office_engines import engine_executable
            executable = engine_executable("msoffice", "writer")
            if not executable:
                raise OSError("Microsoft Word executable is unavailable")
            argv = [executable, str(artifact)]
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
