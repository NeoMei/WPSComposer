from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET


HERE = Path(__file__).resolve().parent
GENERATOR = HERE / "generate_word_style_matrix.py"
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class WordStyleMatrixGeneratorTest(unittest.TestCase):
    def test_generator_produces_exact_deterministic_matrix(self) -> None:
        if not GENERATOR.is_file():
            self.fail("word style matrix generator is missing")
        spec = importlib.util.spec_from_file_location("word_style_matrix_generator", GENERATOR)
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        with tempfile.TemporaryDirectory() as first_raw, tempfile.TemporaryDirectory() as second_raw:
            first = Path(first_raw)
            second = Path(second_raw)
            module.generate(first)
            module.generate(second)

            expected_files = {
                "empty.docx",
                "nonempty.docx",
                "multisection.docx",
                "preflight.json",
                "HASHES.json",
            }
            self.assertEqual({path.name for path in first.iterdir()}, expected_files)
            self.assertEqual({path.name for path in second.iterdir()}, expected_files)
            self.assertEqual(
                {name: sha256(first / name) for name in expected_files},
                {name: sha256(second / name) for name in expected_files},
            )

            preflight = json.loads((first / "preflight.json").read_text(encoding="utf-8"))
            self.assertEqual(preflight["schema"], "wpscomposer-word-style-matrix-v1")
            self.assertEqual(set(preflight["fixtures"]), {"empty", "nonempty", "multisection"})

            empty = preflight["fixtures"]["empty"]
            self.assertEqual(empty["section_count"], 1)
            self.assertEqual(empty["numbering"], [])
            self.assertEqual(empty["custom_style_ids"], [])
            self.assertEqual(empty["all_applied_style_locations"], [{
                "kind": "paragraph",
                "path": "body/p[1]",
                "section": 1,
                "style_id": "Normal",
                "text": "",
            }])
            self.assertNotIn("word/numbering.xml", empty["package_entries"])
            self.assertFalse(any(
                row["type"].endswith("/numbering")
                for row in empty["relationships"]["document"]
            ))

            nonempty = preflight["fixtures"]["nonempty"]
            self.assertEqual(nonempty["section_count"], 1)
            self.assertEqual(nonempty["numbering"], [{
                "abstract_num_id": 4242,
                "level": 0,
                "level_text": "%1.",
                "num_format": "decimal",
                "num_id": 4242,
            }])
            selected = {row["style_id"]: row for row in nonempty["selected_styles"]}
            self.assertEqual(selected["Heading2"]["role"], "applied-modified-builtin")
            self.assertEqual(selected["Heading2"]["auto_redefine"], True)
            self.assertGreater(len(selected["Heading2"]["applied_locations"]), 0)
            self.assertEqual(selected["Heading3"]["role"], "unused-modified-builtin")
            self.assertEqual(selected["Heading3"]["name"], "heading 3")
            self.assertEqual(selected["Heading3"]["style_type"], "paragraph")
            self.assertEqual(selected["Heading3"]["auto_redefine"], True)
            self.assertEqual(selected["Heading3"]["applied_locations"], [])
            self.assertEqual(selected["WPSC_Unused_Custom"]["applied_locations"], [])
            self.assertEqual(
                selected["WPSC_Numbered_Style"]["numbering"],
                {"abstract_num_id": 4242, "level": 0, "num_id": 4242},
            )

            multisection = preflight["fixtures"]["multisection"]
            self.assertEqual(multisection["section_count"], 3)
            all_locations = [
                location
                for row in multisection["selected_styles"]
                for location in row["applied_locations"]
            ]
            self.assertTrue(any(row["section"] == 1 for row in all_locations))
            self.assertTrue(any(row["section"] == 3 for row in all_locations))
            self.assertTrue(any("/tbl[1]/" in row["path"] for row in all_locations))
            self.assertTrue(any(row.get("field_role") == "result" for row in all_locations))

            for fixture_name in ("empty", "nonempty", "multisection"):
                path = first / f"{fixture_name}.docx"
                with zipfile.ZipFile(path) as package:
                    infos = package.infolist()
                    self.assertEqual([info.filename for info in infos], sorted(info.filename for info in infos))
                    self.assertTrue(all(info.date_time == FIXED_ZIP_TIME for info in infos))
                    self.assertTrue(all(info.extra == b"" and info.comment == b"" for info in infos))
                    document = ET.fromstring(package.read("word/document.xml"))
                    self.assertEqual(
                        "".join(node.text or "" for node in document.iter(qn("t"))),
                        "" if fixture_name == "empty" else "".join(
                            row["text"]
                            for row in preflight["fixtures"][fixture_name]["all_applied_style_locations"]
                            if row["kind"] == "paragraph"
                        ),
                    )
                    styles = ET.fromstring(package.read("word/styles.xml"))
                    self.assertTrue(all(qn("val") not in node.attrib for node in styles.iter(qn("spacing"))))


if __name__ == "__main__":
    unittest.main()
