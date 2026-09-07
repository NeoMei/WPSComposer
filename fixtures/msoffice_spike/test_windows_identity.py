"""Portable guard tests; these are not native Word acceptance evidence."""
from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import patch

import windows_word


class IdentityTests(unittest.TestCase):
    def setUp(self):
        self.app = SimpleNamespace(Path="C:/Office", Name="Microsoft Word", Version="16.0", Build="test")

    def test_accepts_single_matching_new_process_without_application_hwnd(self):
        with patch.object(windows_word, "word_processes", return_value={10:"C:/Office/WINWORD.EXE", 20:"C:/Office/WINWORD.EXE"}):
            self.assertEqual(windows_word.application_identity(self.app, {10:"C:/Office/WINWORD.EXE"})["pid"], 20)

    def test_refuses_reused_process(self):
        with patch.object(windows_word, "word_processes", return_value={10:"C:/Office/WINWORD.EXE"}):
            with self.assertRaises(RuntimeError):
                windows_word.application_identity(self.app, {10:"C:/Office/WINWORD.EXE"})

    def test_refuses_ambiguous_new_processes(self):
        with patch.object(windows_word, "word_processes", return_value={10:"C:/Office/WINWORD.EXE", 20:"C:/Office/WINWORD.EXE"}):
            with self.assertRaises(RuntimeError):
                windows_word.application_identity(self.app, {})

    def test_refuses_mismatched_com_path(self):
        with patch.object(windows_word, "word_processes", return_value={10:"C:/Other/WINWORD.EXE"}):
            with self.assertRaises(RuntimeError):
                windows_word.application_identity(self.app, {})

    def test_refuses_non_word_com_name(self):
        self.app.Name = "WPS"
        with patch.object(windows_word, "word_processes", return_value={10:"C:/Office/WINWORD.EXE"}):
            with self.assertRaises(RuntimeError):
                windows_word.application_identity(self.app, {})


if __name__ == "__main__":
    unittest.main()
