from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET


HERE = Path(__file__).resolve().parent
GENERATOR = HERE / "generate_word_story_matrix.py"
PROBE = HERE / "word_story_matrix_probe.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise AssertionError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class WordStoryMatrixTest(unittest.TestCase):
    def test_generator_produces_deterministic_exact_story_topology(self) -> None:
        generator = load(GENERATOR, "word_story_matrix_generator")
        with tempfile.TemporaryDirectory() as first_raw, tempfile.TemporaryDirectory() as second_raw:
            first = Path(first_raw)
            second = Path(second_raw)
            generator.generate(first)
            generator.generate(second)
            expected = {
                "absent.docx",
                "primary.docx",
                "inactive-first-even.docx",
                "linked-multisection.docx",
                "preflight.json",
                "HASHES.json",
            }
            self.assertEqual({path.name for path in first.iterdir()}, expected)
            self.assertEqual(
                {name: sha256(first / name) for name in expected},
                {name: sha256(second / name) for name in expected},
            )
            preflight = json.loads((first / "preflight.json").read_text(encoding="utf-8"))
            self.assertEqual(preflight["schema"], "wpscomposer-word-story-matrix-v1")
            self.assertEqual(
                set(preflight["fixtures"]),
                {"absent", "primary", "inactive-first-even", "linked-multisection"},
            )

            absent = preflight["fixtures"]["absent"]
            self.assertEqual(absent["section_count"], 1)
            self.assertEqual(absent["story_parts"], [])
            self.assertEqual(absent["section_story_references"], [])
            self.assertFalse(absent["even_and_odd_headers"])

            primary = preflight["fixtures"]["primary"]
            self.assertEqual(primary["section_count"], 1)
            self.assertEqual(
                {(row["kind"], row["type"]) for row in primary["section_story_references"]},
                {("header", "default"), ("footer", "default")},
            )
            self.assertTrue(all(row["active"] for row in primary["section_story_references"]))
            self.assertTrue(all(row["ordinary_text"] for row in primary["story_parts"]))
            self.assertTrue(all(row["hidden_text"] for row in primary["story_parts"]))
            self.assertTrue(all(row["fields"] == [{
                "instruction": " PAGE ",
                "locked": True,
                "result": "7",
                "type": "PAGE",
            }] for row in primary["story_parts"]))

            inactive = preflight["fixtures"]["inactive-first-even"]
            self.assertFalse(inactive["even_and_odd_headers"])
            self.assertFalse(inactive["sections"][0]["title_page"])
            self.assertEqual(
                {(row["kind"], row["type"]) for row in inactive["section_story_references"]},
                {("header", "first"), ("header", "even"), ("footer", "first"), ("footer", "even")},
            )
            self.assertTrue(all(not row["active"] for row in inactive["section_story_references"]))

            linked = preflight["fixtures"]["linked-multisection"]
            self.assertEqual(linked["section_count"], 3)
            self.assertTrue(linked["even_and_odd_headers"])
            self.assertTrue(all(row["title_page"] for row in linked["sections"]))
            by_section = {
                section: [row for row in linked["section_story_references"] if row["section"] == section]
                for section in (1, 2, 3)
            }
            self.assertEqual(len(by_section[1]), 6)
            self.assertEqual(by_section[2], [])
            self.assertEqual(len(by_section[3]), 6)
            self.assertEqual(
                {(row["kind"], row["type"]) for row in by_section[1]},
                {(kind, story_type) for kind in ("header", "footer") for story_type in ("default", "first", "even")},
            )
            self.assertTrue(all(row["active"] and not row["linked_to_previous"] for row in by_section[1]))
            self.assertEqual(len(linked["inherited_story_slots"]), 6)
            self.assertTrue(all(row["section"] == 2 and row["linked_to_previous"] for row in linked["inherited_story_slots"]))
            self.assertTrue(all(row["active"] and not row["linked_to_previous"] for row in by_section[3]))

            for fixture_name in ("absent", "primary", "inactive-first-even", "linked-multisection"):
                with zipfile.ZipFile(first / f"{fixture_name}.docx") as package:
                    infos = package.infolist()
                    self.assertEqual([info.filename for info in infos], sorted(info.filename for info in infos))
                    self.assertTrue(all(info.date_time == FIXED_ZIP_TIME for info in infos))
                    self.assertTrue(all(info.extra == b"" and info.comment == b"" for info in infos))
                    rels = ET.fromstring(package.read("word/_rels/document.xml.rels"))
                    targets = {row.attrib["Target"] for row in rels.findall(f"{{{PKG_REL}}}Relationship")}
                    expected_targets = {row["part"] for row in preflight["fixtures"][fixture_name]["story_parts"]}
                    self.assertEqual({target for target in targets if target.startswith(("header", "footer"))}, expected_targets)

    def test_probe_is_opt_in_and_validates_typed_observations(self) -> None:
        probe = load(PROBE, "word_story_matrix_probe")
        self.assertEqual(probe.STATUS_OK, "STORY_MATRIX_DIAGNOSED")
        contract = probe.fixture_contract()
        self.assertEqual(set(contract["story_kinds"]), {
            "primary-header", "first-header", "even-header",
            "primary-footer", "first-footer", "even-footer",
        })
        lines = probe.story_snapshot_lines()
        source = "\n".join(lines)
        self.assertEqual(lines.count("try"), lines.count("end try"))
        self.assertEqual(
            sum(line == "repeat" or line.startswith("repeat with ") for line in lines),
            lines.count("end repeat"),
        )
        self.assertIn("get story range boundDoc story type", source)
        self.assertIn("next story range", source)
        self.assertIn("ordinary-property-error", source)
        self.assertIn("unresolved-property", source)
        self.assertIn("native-missing", source)
        self.assertNotIn("text object of header", source)
        self.assertNotIn("text object of footer", source)
        for term in ("story type", "start of content", "end of content", "content of", "field code", "result range", "locked"):
            self.assertIn(term, source)
        self.assertIn("count sections of qualityStoryRange", source)
        self.assertIn("index of section 1 of qualityStoryRange", source)

        valid = [
            ["story-block", "primary-header"],
            ["story-node", "primary-header", 1],
            ["property-ok", "primary-header", 1, "class", "text", "story range"],
            ["property-ok", "primary-header", 1, "story-type", "constant", "primary header story"],
            ["property-ok", "primary-header", 1, "start", "integer", 12],
            ["property-ok", "primary-header", 1, "end", "integer", 21],
            ["property-ok", "primary-header", 1, "content", "text", "header\r"],
            ["property-ok", "primary-header", 1, "section-count", "integer", 1],
            ["property-ok", "primary-header", 1, "section-first-index", "integer", 1],
            ["field-count", "primary-header", 1, 0],
            ["chain-end", "primary-header", 1, "missing-value"],
        ]
        summary = probe.validate_story_rows(valid, expected_label="primary-header")
        self.assertEqual(summary, {"nodes": 1, "fields": 0, "observations": 9, "ordinary_errors": 0})

        field_rows = valid[:-2] + [
            ["field-count", "primary-header", 1, 1],
            ["field-node", "primary-header", 1, 1],
            ["field-property-ok", "primary-header", 1, 1, "field-type", "constant", "field page"],
            ["field-property-ok", "primary-header", 1, 1, "field-code-content", "text", " PAGE "],
            ["field-property-ok", "primary-header", 1, 1, "field-code-start", "integer", 15],
            ["field-property-ok", "primary-header", 1, 1, "field-code-end", "integer", 21],
            ["field-property-ok", "primary-header", 1, 1, "field-result-content", "text", "7"],
            ["field-property-ok", "primary-header", 1, 1, "field-result-start", "integer", 22],
            ["field-property-ok", "primary-header", 1, 1, "field-result-end", "integer", 23],
            ["field-property-ok", "primary-header", 1, 1, "field-locked", "boolean", True],
            ["chain-end", "primary-header", 1, "missing-value"],
        ]
        self.assertEqual(
            probe.validate_story_rows(field_rows, expected_label="primary-header"),
            {"nodes": 1, "fields": 1, "observations": 17, "ordinary_errors": 0},
        )
        with self.assertRaisesRegex(RuntimeError, "field property coverage"):
            probe.validate_story_rows(field_rows[:-2] + field_rows[-1:], expected_label="primary-header")

        unresolved = [
            ["story-block", "first-footer"],
            ["unresolved-property", "first-footer", 0, "root", "empty-class"],
        ]
        self.assertEqual(
            probe.validate_story_rows(unresolved, expected_label="first-footer"),
            {"nodes": 0, "fields": 0, "observations": 1, "ordinary_errors": 0},
        )
        self.assertEqual(
            probe.validate_story_rows([
                ["story-block", "even-footer"],
                ["unresolved-property", "even-footer", 0, "root", "native-missing"],
            ], expected_label="even-footer")["observations"],
            1,
        )
        with self.assertRaisesRegex(RuntimeError, "empty class cannot be property-ok"):
            probe.validate_story_rows([
                ["story-block", "first-footer"],
                ["property-ok", "first-footer", 0, "root", "", "story range"],
            ], expected_label="first-footer")

        result = subprocess.run(
            [sys.executable, str(PROBE), "--emit-plan"],
            check=True,
            capture_output=True,
            text=True,
        )
        emitted = json.loads(result.stdout)
        self.assertEqual(emitted["status"], "SOURCE_ONLY")
        refused = subprocess.run([sys.executable, str(PROBE)], capture_output=True, text=True)
        self.assertNotEqual(refused.returncode, 0)
        self.assertIn("requires explicit --execute-native", refused.stderr)


if __name__ == "__main__":
    unittest.main()
