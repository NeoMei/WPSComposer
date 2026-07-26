from __future__ import annotations

import pytest

from skills.WPSComposer.scripts import document_api as api
from skills.WPSComposer.scripts.document_api import (
    ALL_OPS,
    INSERT_TYPES,
    PATCH_GRAMMAR,
    PatchError,
    apply_ops,
    apply_patches,
    edit,
    patch_grammar,
    snapshot_to_patches,
    validate_op,
    validate_target,
)


# ---------------------------------------------------------------------------
# Fake composer -- stands in for a COM-backed Writer/Sheet/Slide composer so
# the orchestration layer can be tested without pywin32 or a live WPS host.
# ---------------------------------------------------------------------------

class FakeWriterComposer:
    kind = "writer"

    def __init__(self):
        self.applied = []
        self.saved = False
        self.saved_path = None
        self.closed = False

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False

    def inspect_document(self):
        return {"kind": "writer", "fake": True}

    def apply_format_patch(self, target, **patch):
        # Mirror the real ValueError message shape so enrichment can detect it.
        if target == "boom":
            raise ValueError("Unsupported Writer target: boom")
        if target == "explode":
            raise RuntimeError("COM host vanished")
        accepted = sorted(patch.keys())
        self.applied.append((target, patch))
        return {"accepted": accepted, "rejected": []}

    def apply_structural_op(self, op):
        # Fake structural dispatcher: records the op and returns a fake path.
        verb = op.get("op")
        if verb == "remove" and op.get("target") == "paragraph:404":
            raise ValueError("Unsupported Writer target: paragraph:404")
        if verb == "insert":
            etype = op.get("type")
            self.applied.append(("insert", op))
            return {"type": etype, "path": "paragraph:999"}
        self.applied.append((verb, op))
        return {"path": op.get("target")}

    def save(self, path):
        self.saved = True
        self.saved_path = path
        return path

    def save_current(self):
        self.saved = True
        self.saved_path = "<active>"
        return self.saved_path

    def export_pdf(self, path):
        return path

    def close(self, save_changes=False):
        self.closed = True


# ---------------------------------------------------------------------------
# validate_target / patch_grammar
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("kind,target,element", [
    ("writer", "selection", "selection"),
    ("writer", "paragraph:3", "paragraph"),
    ("writer", "range:0-10", "range"),
    ("writer", "table:1/cell:2,3", "table_cell"),
    ("writer", "shape:5", "shape"),
    ("writer", "section:1", "section"),
    ("sheet", "sheet:2/cell:$A$1", "cell"),
    ("sheet", "sheet:1/range:A1:C20", "range"),
    ("slide", "slide:1/shape:2/paragraph:3/run:1", "run"),
    ("slide", "presentation", "presentation"),
    ("word", "paragraph:1", "paragraph"),      # alias accepted
    ("excel", "sheet:1/cell:A1", "cell"),      # alias accepted
])
def test_validate_target_accepts_valid_forms(kind, target, element):
    info = validate_target(target, kind)
    assert info["valid"] is True
    assert info["element"] == element
    assert info["kind"] in {"writer", "sheet", "slide"}


def test_validate_target_rejects_unknown_kind():
    info = validate_target("paragraph:1", "nope")
    assert info["valid"] is False
    assert info["error"]["code"] == "unsupported_kind"
    assert "writer" in info["error"]["valid_forms"]


def test_validate_target_gives_suggestion_on_bad_target():
    info = validate_target("paragraf:3", "writer")
    assert info["valid"] is False
    err = info["error"]
    assert err["code"] == "invalid_target"
    assert "paragraph:N" in err["valid_forms"]
    assert err["closest"] == "paragraph:N"


def test_patch_grammar_returns_all_kinds():
    grammar = patch_grammar()
    assert set(grammar) == {"writer", "sheet", "slide"}
    for entries in grammar.values():
        assert all({"form", "element", "description"} <= set(e) for e in entries)


def test_patch_grammar_scoped_to_one_kind():
    entries = patch_grammar("writer")
    forms = {e["form"] for e in entries}
    assert "paragraph:N" in forms
    with pytest.raises(ValueError):
        patch_grammar("bogus")


def test_patch_grammar_matches_inspect_forms():
    # The grammar advertised to agents must match the regex table.
    assert {k for k in PATCH_GRAMMAR} == {"writer", "sheet", "slide"}


# ---------------------------------------------------------------------------
# apply_patches -- structured reports + atomic semantics
# ---------------------------------------------------------------------------

def test_apply_patches_returns_structured_reports():
    composer = FakeWriterComposer()
    reports = apply_patches(
        composer,
        [{"target": "paragraph:1", "font": {"size": 12}},
         {"target": "paragraph:2", "font": {"bold": True}}],
    )
    assert all(r["ok"] for r in reports)
    assert reports[0]["accepted"] == ["font"]
    assert composer.applied[0] == ("paragraph:1", {"font": {"size": 12}})


def test_apply_patches_missing_target_is_atomic_failure():
    composer = FakeWriterComposer()
    with pytest.raises(PatchError) as exc_info:
        apply_patches(composer, [{"font": {}}])
    report = exc_info.value.reports[0]
    assert report["ok"] is False
    assert report["error"]["code"] == "missing_target"


def test_apply_patches_enriches_invalid_target_with_suggestion():
    composer = FakeWriterComposer()
    with pytest.raises(PatchError) as exc_info:
        apply_patches(composer, [{"target": "boom", "font": {}}])
    err = exc_info.value.reports[0]["error"]
    assert err["code"] == "invalid_target"
    assert "paragraph:N" in err["valid_forms"]


def test_apply_patches_distinguishes_apply_failure_code():
    composer = FakeWriterComposer()
    with pytest.raises(PatchError) as exc_info:
        apply_patches(composer, [{"target": "explode"}])
    assert exc_info.value.reports[0]["error"]["code"] == "apply_failed"


def test_apply_patches_best_effort_keeps_running():
    composer = FakeWriterComposer()
    reports = apply_patches(
        composer,
        [{"target": "paragraph:1", "font": {"size": 12}},
         {"target": "boom"},
         {"target": "paragraph:2", "font": {"bold": True}}],
        atomic=False,
    )
    assert reports[0]["ok"] is True
    assert reports[1]["ok"] is False
    assert reports[2]["ok"] is True          # kept going past the failure
    assert composer.applied[-1][0] == "paragraph:2"


def test_apply_patches_stop_on_error_alias_maps_to_atomic():
    # Backward-compat: stop_on_error=False == atomic=False
    composer = FakeWriterComposer()
    reports = apply_patches(
        composer, [{"target": "boom"}], stop_on_error=False
    )
    assert reports[0]["ok"] is False


# ---------------------------------------------------------------------------
# edit -- atomic guarantee: no save on failure, structured result
# ---------------------------------------------------------------------------

def test_edit_returns_success_result_and_saves():
    composer = FakeWriterComposer()
    # edit() opens its own composer via open_document; patch that factory.
    captured = {}

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        captured["path"] = path
        return composer

    monkey = _Monkey(api, "open_document", fake_open)
    with monkey:
        result = api.edit(
            "report.docx",
            patches=[{"target": "paragraph:1", "font": {"size": 12}}],
            output="out.docx",
        )
    assert result["ok"] is True
    assert result["saved"] is True
    assert result["saved_path"] == "out.docx"
    assert result["errors"] == []
    assert composer.saved is True
    assert composer.closed is True


def test_edit_atomic_blocks_save_on_failure():
    composer = FakeWriterComposer()

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        return composer

    with _Monkey(api, "open_document", fake_open):
        result = api.edit(
            "report.docx",
            patches=[{"target": "boom"}],
            output="out.docx",
        )
    assert result["ok"] is False
    assert result["saved"] is False
    assert result["saved_path"] is None
    assert result["errors"][0]["error"]["code"] == "invalid_target"
    assert composer.saved is False          # the atomic guarantee
    assert composer.closed is True


def test_edit_best_effort_saves_partial():
    composer = FakeWriterComposer()

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        return composer

    with _Monkey(api, "open_document", fake_open):
        result = api.edit(
            "report.docx",
            patches=[{"target": "paragraph:1", "font": {"size": 12}},
                     {"target": "boom"}],
            output="out.docx",
            atomic=False,
        )
    assert result["ok"] is True             # not raised, best-effort
    assert result["saved"] is True
    assert composer.saved is True
    assert any(not r["ok"] for r in result["patches"])


def test_edit_raise_on_error_propagates_patch_error():
    composer = FakeWriterComposer()

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        return composer

    with _Monkey(api, "open_document", fake_open):
        with pytest.raises(PatchError):
            api.edit(
                "report.docx",
                patches=[{"target": "boom"}],
                output="out.docx",
                raise_on_error=True,
            )
    assert composer.saved is False


# ---------------------------------------------------------------------------
# Tiny context-manager for monkeypatching module attributes (no pytest plugin)
# ---------------------------------------------------------------------------

class _Monkey:
    def __init__(self, module, name, value):
        self.module = module
        self.name = name
        self.value = value
        self._original = getattr(module, name)

    def __enter__(self):
        setattr(self.module, self.name, self.value)
        return self

    def __exit__(self, *_):
        setattr(self.module, self.name, self._original)
        return False


# ---------------------------------------------------------------------------
# Stable-id grammar forms (@paraId / @id / @name)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("target,kind,element", [
    ("paragraph:@paraId=1A2B3C4D", "writer", "paragraph"),
    ("paragraph:@paraId=abcdef12", "writer", "paragraph"),   # lowercase ok
    ("slide:1/shape:@id=123", "slide", "shape"),
    ("slide:1/shape:@name=Title 1", "slide", "shape"),       # spaces in name
    ("sheet:2/shape:@id=5", "sheet", "shape"),
    ("sheet:2/shape:@name=My Logo!", "sheet", "shape"),
])
def test_validate_target_accepts_stable_id_forms(target, kind, element):
    info = validate_target(target, kind)
    assert info["valid"] is True
    assert info["element"] == element


def test_validate_target_rejects_malformed_paraid():
    # non-hex char must not match the paraId form
    assert validate_target("paragraph:@paraId=ZZ12", "writer")["valid"] is False


def test_patch_grammar_advertises_stable_forms():
    writer_forms = {e["form"] for e in patch_grammar("writer")}
    slide_forms = {e["form"] for e in patch_grammar("slide")}
    sheet_forms = {e["form"] for e in patch_grammar("sheet")}
    assert "paragraph:@paraId=HEX" in writer_forms
    assert "slide:N/shape:@id=N" in slide_forms
    assert "slide:N/shape:@name=NAME" in slide_forms
    assert "sheet:N/shape:@id=N" in sheet_forms
    assert "sheet:N/shape:@name=NAME" in sheet_forms


# ---------------------------------------------------------------------------
# Word w14:paraId extraction (pure, no COM)
# ---------------------------------------------------------------------------

from skills.WPSComposer.scripts.writer import _extract_paraids

_SAMPLE_DOCUMENT_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="0A1B2C3D"><w:r><w:t>first</w:t></w:r></w:p>
    <w:p><w:r><w:t>second, no paraId</w:t></w:r></w:p>
    <w:tbl>
      <w:tr><w:tc>
        <w:p w14:paraId="FEDCBA98"><w:r><w:t>inside a table cell</w:t></w:r></w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:p w14:paraId="11223344"><w:r><w:t>fourth</w:t></w:r></w:p>
  </w:body>
</w:document>
"""


def test_extract_paraids_returns_document_order_with_none_for_missing():
    # Order must match doc.Paragraphs(1..N): body para, body para, table-cell
    # para, trailing body para. Missing w14:paraId -> None.
    assert _extract_paraids(_SAMPLE_DOCUMENT_XML) == [
        "0A1B2C3D", None, "FEDCBA98", "11223344",
    ]


def test_extract_paraids_empty_body():
    empty = (b'<w:document xmlns:w="http://schemas.openxmlformats.org/'
             b'wordprocessingml/2006/main"><w:body/></w:document>')
    assert _extract_paraids(empty) == []


def test_read_paraids_from_docx_handles_missing_file(tmp_path):
    from skills.WPSComposer.scripts.writer import read_paraids_from_docx
    # Non-existent path -> [] (no exception)
    assert read_paraids_from_docx(tmp_path / "does-not-exist.docx") == []


def test_read_paraids_from_docx_handles_bad_zip(tmp_path):
    import zipfile
    from skills.WPSComposer.scripts.writer import read_paraids_from_docx
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"not a zip file")
    # Corrupt / non-docx -> [] (no exception); the wrapper swallows parse errors
    assert read_paraids_from_docx(bad) == []


def test_read_paraids_from_docx_round_trips_real_docx(tmp_path):
    import zipfile
    from skills.WPSComposer.scripts.writer import read_paraids_from_docx
    docx = tmp_path / "sample.docx"
    with zipfile.ZipFile(docx, "w") as archive:
        archive.writestr("word/document.xml", _SAMPLE_DOCUMENT_XML)
    assert read_paraids_from_docx(docx) == [
        "0A1B2C3D", None, "FEDCBA98", "11223344",
    ]


# ---------------------------------------------------------------------------
# snapshot_to_patches (dump -> replay)
# ---------------------------------------------------------------------------

def test_snapshot_to_patches_emits_one_patch_per_styled_element():
    snapshot = {
        "kind": "writer",
        "paragraphs": [
            {"id": "paragraph:@paraId=1A2B3C4D", "index": 1,
             "font": {"size": 12, "bold": True},
             "paragraph": {"alignment": 3}},
            {"id": "paragraph:2", "index": 2, "font": {}},   # empty -> skip
        ],
        "tables": [
            {"id": "table:1", "cells": [
                {"id": "table:1/cell:1,1", "font": {"name": "Arial"},
                 "fill": {"color": "#FFF2CC"}},
            ]},
        ],
    }
    patches = snapshot_to_patches(snapshot)
    targets = [p["target"] for p in patches]
    assert targets == ["paragraph:@paraId=1A2B3C4D", "table:1/cell:1,1"]
    assert patches[0]["font"] == {"size": 12, "bold": True}
    assert patches[0]["paragraph"] == {"alignment": 3}
    assert patches[1]["fill"] == {"color": "#FFF2CC"}


def test_snapshot_to_patches_respects_dimension_filter():
    snapshot = {
        "kind": "writer",
        "paragraphs": [
            {"id": "paragraph:1", "font": {"size": 12},
             "paragraph": {"alignment": 3}, "fill": {"color": "#000000"}},
        ],
    }
    patches = snapshot_to_patches(snapshot, dimensions=("font",))
    assert patches == [{"target": "paragraph:1", "font": {"size": 12}}]


def test_snapshot_to_patches_walks_nested_slide_shapes():
    snapshot = {
        "kind": "slide",
        "slides": [
            {"id": "slide:1", "shapes": [
                {"id": "slide:1/shape:@id=7", "font": {"size": 18}},
                {"id": "slide:1/shape:2", "fill": {"color": "#FF0000"}},
            ]},
        ],
    }
    patches = snapshot_to_patches(snapshot)
    assert {p["target"] for p in patches} == {
        "slide:1/shape:@id=7", "slide:1/shape:2",
    }


def test_snapshot_to_patches_skips_invalid_ids():
    # An id that matches no grammar form for the snapshot's kind is skipped.
    snapshot = {
        "kind": "writer",
        "paragraphs": [{"id": "whoami:5", "font": {"size": 12}}],
    }
    assert snapshot_to_patches(snapshot) == []


# ---------------------------------------------------------------------------
# validate_op + apply_ops (structural verbs)
# ---------------------------------------------------------------------------

def test_validate_op_accepts_set():
    info = validate_op({"op": "set", "target": "paragraph:1"}, "writer")
    assert info["valid"] is True and info["verb"] == "set"


def test_validate_op_rejects_unknown_verb():
    info = validate_op({"op": "frobnicate", "target": "paragraph:1"}, "writer")
    assert info["valid"] is False
    assert info["error"]["code"] == "unknown_verb"
    assert "set" in info["error"]["valid_verbs"]


def test_validate_op_insert_requires_type():
    info = validate_op({"op": "insert", "parent": "body"}, "writer")
    assert info["valid"] is False
    assert info["error"]["code"] == "missing_type"


def test_validate_op_insert_rejects_unsupported_type_per_kind():
    info = validate_op({"op": "insert", "type": "slide"}, "writer")
    assert info["valid"] is False
    assert info["error"]["code"] == "unsupported_type"
    assert "paragraph" in info["error"]["valid_types"]


@pytest.mark.parametrize("verb", ["remove", "move", "clone"])
def test_validate_op_structural_verbs_require_target(verb):
    info = validate_op({"op": verb}, "writer")
    assert info["valid"] is False
    assert info["error"]["code"] == "missing_target"


def test_validate_op_position_accepts_after_anchor():
    info = validate_op(
        {"op": "insert", "type": "paragraph",
         "position": {"after": "paragraph:2"}},
        "writer",
    )
    assert info["valid"] is True
    assert info["position"] == {"after": "paragraph:2"}


def test_validate_op_position_accepts_slide_shape_destination():
    # Slide shape move/clone use {"slide": N} -- must NOT be rejected.
    info = validate_op(
        {"op": "move", "target": "slide:1/shape:2", "to": {"slide": 3}},
        "slide",
    )
    assert info["valid"] is True


def test_validate_op_position_accepts_int_sheet_anchors():
    # Sheet sheet move/clone use {"after": N} / {"before": N} with int sheet
    # numbers -- must NOT crash validate_target on an int.
    for spec in ({"after": 2}, {"before": 3}):
        info = validate_op(
            {"op": "move", "target": "sheet:1", "to": spec}, "sheet"
        )
        assert info["valid"] is True, (spec, info)


def test_validate_op_position_rejects_bad_anchor():
    info = validate_op(
        {"op": "insert", "type": "paragraph",
         "position": {"after": "whoami:9"}},
        "writer",
    )
    assert info["valid"] is False
    assert info["error"]["code"] == "invalid_anchor"


@pytest.mark.parametrize("kind,etype", [
    ("writer", "paragraph"), ("writer", "heading"), ("writer", "page_break"),
    ("writer", "table"), ("writer", "image"), ("writer", "textbox"),
    ("slide", "slide"), ("slide", "textbox"), ("slide", "image"),
    ("sheet", "row"), ("sheet", "column"), ("sheet", "sheet"),
])
def test_validate_op_accepts_all_documented_insert_types(kind, etype):
    # parent chosen so position anchors (if any) validate; default 'end' is fine
    info = validate_op({"op": "insert", "type": etype}, kind)
    assert info["valid"] is True, (kind, etype, info)


def test_apply_ops_dispatches_clone_verb():
    composer = FakeWriterComposer()
    reports = apply_ops(composer, [
        {"op": "clone", "target": "paragraph:2", "to": "end"},
    ])
    assert reports[0]["op"] == "clone"
    assert reports[0]["ok"] is True
    # fake composer records the dispatched verb
    assert composer.applied[-1][0] == "clone"


def test_validate_op_ignores_unknown_extra_fields_like_axis():
    # Sheet remove/move accept an `axis` hint; validate_op must not reject it.
    info = validate_op(
        {"op": "remove", "target": "sheet:1/cell:C1", "axis": "column"},
        "sheet",
    )
    assert info["valid"] is True


def test_apply_ops_forwards_axis_field_to_structural_handler():
    composer = FakeWriterComposer()
    apply_ops(composer, [
        {"op": "remove", "target": "paragraph:2", "axis": "column"},
    ])
    # the fake records the full op dict for structural verbs
    recorded_op = composer.applied[-1][1]
    assert recorded_op.get("axis") == "column"


def test_apply_ops_runs_mixed_set_and_structural():
    composer = FakeWriterComposer()
    reports = apply_ops(composer, [
        {"op": "set", "target": "paragraph:1", "font": {"size": 12}},
        {"op": "insert", "parent": "body", "type": "paragraph",
         "props": {"text": "new"}},
        {"op": "remove", "target": "paragraph:2"},
    ])
    assert [r["op"] for r in reports] == ["set", "insert", "remove"]
    assert all(r["ok"] for r in reports)
    assert reports[1]["path"] == "paragraph:999"
    # set went to apply_format_patch, structural went to apply_structural_op
    assert composer.applied[0] == ("paragraph:1", {"font": {"size": 12}})
    assert composer.applied[1][0] == "insert"


def test_apply_ops_structural_failure_is_atomic():
    composer = FakeWriterComposer()
    with pytest.raises(PatchError) as exc_info:
        apply_ops(composer, [
            {"op": "insert", "type": "paragraph"},
            {"op": "remove", "target": "paragraph:404"},  # fake raises
        ])
    reports = exc_info.value.reports
    assert reports[0]["ok"] is True
    assert reports[1]["ok"] is False
    assert reports[1]["error"]["code"] == "invalid_target"


def test_apply_ops_best_effort_continues_past_structural_failure():
    composer = FakeWriterComposer()
    reports = apply_ops(composer, [
        {"op": "insert", "type": "paragraph"},
        {"op": "remove", "target": "paragraph:404"},
        {"op": "insert", "type": "heading"},
    ], atomic=False)
    assert reports[1]["ok"] is False
    assert reports[2]["ok"] is True


def test_apply_patches_still_works_as_set_wrapper():
    composer = FakeWriterComposer()
    reports = apply_patches(
        composer, [{"target": "paragraph:1", "font": {"size": 12}}]
    )
    assert reports[0]["op"] == "set"
    assert reports[0]["ok"] is True


def test_edit_accepts_ops_and_patches_combined():
    composer = FakeWriterComposer()

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        return composer

    with _Monkey(api, "open_document", fake_open):
        result = api.edit(
            "report.docx",
            output="out.docx",
            patches=[{"target": "paragraph:1", "font": {"size": 12}}],
            ops=[{"op": "insert", "parent": "body", "type": "paragraph",
                  "props": {"text": "x"}}],
        )
    assert result["ok"] is True
    assert result["saved"] is True
    # patches run first, then ops
    ops = result["ops"]
    assert [o["op"] for o in ops] == ["set", "insert"]


def test_edit_atomic_blocks_save_on_structural_failure():
    composer = FakeWriterComposer()

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        return composer

    with _Monkey(api, "open_document", fake_open):
        result = api.edit(
            "report.docx",
            output="out.docx",
            ops=[{"op": "insert", "type": "paragraph"},
                 {"op": "remove", "target": "paragraph:404"}],
        )
    assert result["ok"] is False
    assert result["saved"] is False
    assert composer.saved is False
    assert result["errors"][0]["error"]["code"] == "invalid_target"

