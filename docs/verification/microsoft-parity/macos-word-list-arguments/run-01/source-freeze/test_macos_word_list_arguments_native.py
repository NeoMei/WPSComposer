from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock
from zipfile import ZipFile


HERE = Path(__file__).resolve().parent
RUNNER = HERE / "macos_word_list_arguments_native.py"


def load_runner():
    spec = importlib.util.spec_from_file_location("word_list_native", RUNNER)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class SourceOnlyListFixtureTest(unittest.TestCase):
    def test_plan_contains_every_argument_case_and_body_boundaries(self):
        module = load_runner()
        build = module.build_case()
        rows = []
        for operation in build.plan.operations:
            args = dict(operation.args)
            if "items" in args:
                args["items"] = list(args["items"])
            rows.append((operation.op, args))
        lists = [args for op, args in rows if op == "writer.add_list"]

        self.assertEqual(
            lists,
            [
                {"items": ["DEFAULT ITEM"], "ordered": False},
                {"items": ["CUSTOM ITEM"], "ordered": False, "glyph": "→", "indent": 30},
                {"items": ["EMPTY GLYPH ITEM"], "ordered": False, "glyph": "", "indent": 27.5},
                {"items": ["ORDERED ITEM"], "ordered": True, "glyph": 'ignored"\\unsafe', "indent": 33},
            ],
        )
        body_texts = [args["text"] for op, args in rows if op == "writer.add_paragraph"]
        self.assertIn("BODY AFTER DEFAULT", body_texts)
        self.assertIn("BODY AFTER ORDERED", body_texts)

    def test_compiled_source_has_literal_prefixes_resets_and_no_native_list_template(self):
        module = load_runner()
        source = module.compile_case(Path("/preflight/owned.docx")).source

        self.assertIn('"\u2022" & tab & "DEFAULT ITEM" & return', source)
        self.assertIn('"\u2192" & tab & "CUSTOM ITEM" & return', source)
        self.assertIn('"" & tab & "EMPTY GLYPH ITEM" & return', source)
        self.assertIn('"1." & tab & "ORDERED ITEM" & return', source)
        self.assertNotIn("ignored", source)
        self.assertEqual(source.count("set style of r to style list paragraph"), 4)
        self.assertEqual(source.count("reset font object of r"), 4)
        self.assertEqual(source.count("set style of trailingRange to style normal"), 4)
        self.assertNotIn("apply bullet default", source)
        self.assertNotIn("apply number default", source)

    def test_ooxml_text_reads_run_tabs_but_not_paragraph_property_tabs(self):
        module = load_runner()
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        document = f'''<w:document xmlns:w="{namespace}"><w:body><w:p>
          <w:pPr><w:tabs><w:tab w:val="left" w:pos="480"/></w:tabs></w:pPr>
          <w:r><w:t>•</w:t><w:tab/><w:t>ITEM</w:t></w:r>
        </w:p><w:sectPr/></w:body></w:document>'''
        styles = f'''<w:styles xmlns:w="{namespace}"/>'''
        with tempfile.TemporaryDirectory() as dirname:
            path = Path(dirname) / "sample.docx"
            with ZipFile(path, "w") as package:
                package.writestr("word/document.xml", document)
                package.writestr("word/styles.xml", styles)
            rows, _ = module._paragraphs(path)
        self.assertEqual(rows[0]["text"], "•\tITEM")

    def test_native_readback_requires_all_list_and_following_body_formats(self):
        module = load_runner()
        good = [
            [1, "•\tDEFAULT ITEM\r", "List Paragraph", 24, -24],
            [2, "BODY AFTER DEFAULT\r", "Body Text", 0, 24],
            [3, "→\tCUSTOM ITEM\r", "List Paragraph", 30, -30],
            [4, "\tEMPTY GLYPH ITEM\r", "List Paragraph", 27.5, -27.5],
            [5, "1.\tORDERED ITEM\r", "List Paragraph", 33, -33],
            [6, "BODY AFTER ORDERED\r", "Body Text", 0, 24],
        ]
        self.assertTrue(module.validate_readback(good))
        self.assertFalse(module.validate_readback(good[:-1]))
        wrong_indent = [list(row) for row in good]
        wrong_indent[2][3] = 24
        self.assertFalse(module.validate_readback(wrong_indent))

    def test_native_execution_is_opt_in_and_source_freeze_is_required(self):
        module = load_runner()
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "evidence"
            self.assertEqual(module.main(["--output", str(output)]), 2)
            self.assertFalse(output.exists())

            freeze = Path(temp) / "freeze.json"
            freeze.write_text(json.dumps({"schema": 1, "source_hashes": {}}))
            with self.assertRaisesRegex(RuntimeError, "source freeze"):
                module.verify_source_freeze(freeze)

    def test_runner_orchestration_with_fake_owned_runtime(self):
        module = load_runner()

        class FakeAdapter:
            raise_on_close = False

            def __init__(self, build):
                self.build = build
                self.staging_root = None

            def execute(self, build, directives, deadline):
                self.staging_root = temp / "adapter-runtime"
                self.staging_root.mkdir(exist_ok=True)
                staged = self.staging_root / "owned.docx"
                staged.write_bytes(b"fake-docx")
                return SimpleNamespace(staged_artifact=str(staged), issues=())

            def export_pdf(self, staged, deadline):
                target = self.staging_root / "owned.pdf"
                target.write_bytes(b"fake-pdf")
                return target

            def publish(self, staged, output, overwrite, deadline):
                output.write_bytes(Path(staged).read_bytes())
                return output

            def close(self):
                if self.raise_on_close:
                    raise RuntimeError("fake close failure")
                return None

        class FakeSession:
            _closed = False
            staging_root = None

            @classmethod
            def open_document(cls, path, *, read_only, visible):
                self = cls()
                self._closed = False
                return self

            def __enter__(self):
                return self

            def __exit__(self, *args):
                self._closed = True

            def _execute(self, lines):
                return [
                    [1, "•\tDEFAULT ITEM\r", "List Paragraph", 24, -24],
                    [2, "BODY AFTER DEFAULT\r", "Body Text", 0, 24],
                    [3, "→\tCUSTOM ITEM\r", "List Paragraph", 30, -30],
                    [4, "\tEMPTY GLYPH ITEM\r", "List Paragraph", 27.5, -27.5],
                    [5, "1.\tORDERED ITEM\r", "List Paragraph", 33, -33],
                    [6, "BODY AFTER ORDERED\r", "Body Text", 0, 24],
                ]

        with tempfile.TemporaryDirectory() as dirname:
            temp = Path(dirname)
            freeze = temp / "freeze.json"
            hashes = {source.relative_to(module.ROOT).as_posix(): module.sha256(source)
                      for source in module.OWNED_SOURCES}
            freeze.write_text(json.dumps({"schema": 1, "source_hashes": hashes}))
            output = temp / "evidence"
            with mock.patch.object(module, "inspect_docx", return_value={"docx_contract": True}):
                report = module.run(
                    output, freeze, 30,
                    adapter_factory=FakeAdapter,
                    session_type=FakeSession,
                    inventory_reader=lambda *_: [],
                    pdf_inspector=lambda _: {"pdf_contract": True},
                )

            self.assertEqual(report["status"], "PASS")
            self.assertTrue(report["checks"]["read_only_reopen_preserves_bytes"])
            self.assertTrue(report["checks"]["document_inventory_preserved"])
            self.assertEqual(set(report["artifact_hashes"]), {"list-arguments.docx", "list-arguments.pdf"})

            FakeAdapter.raise_on_close = True
            second_output = temp / "close-failure-evidence"
            with mock.patch.object(module, "inspect_docx", return_value={"docx_contract": True}):
                failed = module.run(
                    second_output, freeze, 30,
                    adapter_factory=FakeAdapter,
                    session_type=FakeSession,
                    inventory_reader=lambda *_: [],
                    pdf_inspector=lambda _: {"pdf_contract": True},
                )
            self.assertEqual(failed["status"], "FAIL")
            self.assertEqual(failed["cleanup_error"]["message"], "fake close failure")
            self.assertTrue(failed["checks"]["document_inventory_preserved"])
            self.assertTrue((second_output / "report.json").is_file())


if __name__ == "__main__":
    unittest.main()
