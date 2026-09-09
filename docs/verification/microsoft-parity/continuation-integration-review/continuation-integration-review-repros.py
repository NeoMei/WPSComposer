"""Pure integration repros for continuation-integration-review.md.

No Office process is launched.  The real MacWordSession transport envelope and
recovery controller are used while only subprocess.run is replaced.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory

import pytest

from skills.WPSComposer.scripts.msoffice import macos_word_session as session_api
from skills.WPSComposer.scripts.writer import NativeWriterObjectError


def state(end: int = 6):
    prefix = "target"
    return [
        [
            "checkpoint-state", 1, end, end + 1, 1, 0, end + 1,
            hashlib.sha256(prefix.encode()).hexdigest(), 0, 6, prefix,
            hashlib.sha256((prefix + "\r").encode()).hexdigest(),
        ],
        ["objects", 0, 0, 0, 0, 0],
        ["checkpoint-end"],
    ]


def test_ordinary_native_error_during_destructive_rollback_must_quarantine():
    with TemporaryDirectory() as directory:
        root = Path(directory)
        session = session_api.MacWordSession()
        session.staging_root = root
        session._owns_doc = True
        session._bound_path = str(root / "owned.docx")
        calls = []
        replies = [
            (0, state(), ""),
            (0, state(10), ""),
            # The rollback script can already have deleted fields/tables/text.
            # A normal AppleScript error does not prove that its prefix was undone.
            (1, None, "execution error: WPSC_CHECKPOINT_POSTCONDITION_FAILED (-2700)"),
            (0, [["later-event-ran"]], ""),
        ]

        def fake_run(command, **kwargs):
            calls.append(Path(command[1]).read_text(encoding="utf-8"))
            returncode, rows, stderr = replies.pop(0)
            stdout = "" if rows is None else json.dumps(
                [session_api._COMPLETION_MARKER, rows]
            )
            return session_api.subprocess.CompletedProcess(
                command, returncode, stdout, stderr
            )

        original_run = session_api.subprocess.run
        session_api.subprocess.run = fake_run
        try:
            checkpoint = session.degradation_checkpoint()
            try:
                session.rollback_degradation_checkpoint(checkpoint)
            except NativeWriterObjectError as error:
                assert error.code == "LOCAL_MUTATION_ROLLBACK_FAILED"
            else:
                raise AssertionError("rollback unexpectedly succeeded")

            assert "set content of rollbackRange to \"\"" in calls[2]
            assert session._retain_evidence is True

            # RED on the frozen source: the destructive failure remains usable.
            assert session._quarantined is True
            with pytest.raises(session_api.NativeWordError) as later:
                session._execute(['set nativeRows to {{"later-event-ran"}}'])
            assert later.value.code == "NATIVE_WORD_QUARANTINED"
            assert len(calls) == 3
        finally:
            session_api.subprocess.run = original_run


if __name__ == "__main__":
    test_ordinary_native_error_during_destructive_rollback_must_quarantine()
