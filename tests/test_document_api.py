from __future__ import annotations

from pathlib import Path
import zipfile

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
from tests._pdf_fixture import write_minimal_pdf


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
        if target == "rejected":
            return {"accepted": [], "rejected": ["font.size"]}
        accepted = sorted(patch.keys())
        self.applied.append((target, patch))
        return {"accepted": accepted, "rejected": []}

    def apply_structural_op(self, op):
        # Fake structural dispatcher: records the op and returns a fake path.
        verb = op.get("op")
        if verb == "remove" and op.get("target") == "paragraph:404":
            raise ValueError("Unsupported Writer target: paragraph:404")
        if verb == "remove" and op.get("target") == "paragraph:405":
            return {"accepted": [], "rejected": ["remove"]}
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


class PackageWriterComposer(FakeWriterComposer):
    def __init__(self, *, corrupt=False, fail=False):
        super().__init__()
        self.corrupt = corrupt
        self.fail = fail
        self.save_calls = []

    def save(self, path):
        target = Path(path)
        self.save_calls.append(target.resolve())
        target.parent.mkdir(parents=True, exist_ok=True)
        if self.fail:
            target.write_bytes(b"partial")
            raise OSError("simulated save failure")
        if self.corrupt:
            target.write_bytes(b"not-an-ooxml-package")
        else:
            with zipfile.ZipFile(target, "w") as package:
                package.writestr("[Content_Types].xml", "<Types />")
                package.writestr(
                    "word/document.xml",
                    "<document><body>edited</body></document>",
                )
        self.saved = True
        return str(target)


class PackageAndPdfWriterComposer(PackageWriterComposer):
    def __init__(self, *, pdf_failure=False):
        super().__init__()
        self.pdf_failure = pdf_failure
        self.export_calls = []

    def export_pdf(self, path):
        target = Path(path)
        self.export_calls.append(target.resolve())
        target.parent.mkdir(parents=True, exist_ok=True)
        if self.pdf_failure:
            target.write_bytes(b"PARTIAL")
            raise OSError("simulated PDF export failure")
        write_minimal_pdf(target, b"edited-pdf")
        return str(target)


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


def test_apply_patches_rejected_report_is_atomic_failure():
    composer = FakeWriterComposer()

    with pytest.raises(PatchError) as exc_info:
        apply_patches(
            composer,
            [{"target": "rejected", "font": {"size": 12}}],
        )

    report = exc_info.value.reports[0]
    assert report["ok"] is False
    assert report["rejected"] == ["font.size"]


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
    assert result["ok"] is False
    assert result["saved"] is True
    assert composer.saved is True
    assert any(not r["ok"] for r in result["patches"])
    assert result["errors"] == [r for r in result["patches"] if not r["ok"]]


def test_edit_atomic_rejected_report_blocks_save():
    composer = FakeWriterComposer()

    def fake_open(path, *, kind=None, read_only=False, visible=False):
        return composer

    with _Monkey(api, "open_document", fake_open):
        result = api.edit(
            "report.docx",
            patches=[{"target": "rejected", "font": {"size": 12}}],
            output="out.docx",
        )

    assert result["ok"] is False
    assert result["saved"] is False
    assert result["errors"][0]["rejected"] == ["font.size"]
    assert composer.saved is False


def test_edit_attached_atomic_multi_op_rejected_before_first_mutation():
    composer = FakeWriterComposer()

    with _Monkey(api, "attach_active", lambda kind=None: composer):
        result = api.edit(
            patches=[
                {"target": "paragraph:1", "font": {"bold": True}},
                {"target": "paragraph:2", "font": {"italic": True}},
            ],
            atomic=True,
        )

    assert result["ok"] is False
    assert result["saved"] is False
    assert result["errors"][0]["error"]["code"] == "atomic_attached_batch_unsupported"
    assert composer.applied == []
    assert composer.saved is False


def test_edit_attached_atomic_compound_patch_rejected_before_partial_mutation():
    class CompoundPatchComposer(FakeWriterComposer):
        def __init__(self):
            super().__init__()
            self.live = {}

        def apply_format_patch(self, target, **patch):
            self.live["bold"] = patch["font"]["bold"]
            return {
                "accepted": ["font.bold"],
                "rejected": ["font.unsupported"],
            }

    composer = CompoundPatchComposer()

    with _Monkey(api, "attach_active", lambda kind=None: composer):
        result = api.edit(
            patches=[{
                "target": "paragraph:1",
                "font": {"bold": True, "unsupported": True},
            }],
            atomic=True,
        )

    assert result["ok"] is False
    assert result["saved"] is False
    assert result["errors"][0]["error"]["code"] == "atomic_attached_batch_unsupported"
    assert composer.live == {}
    assert composer.applied == []


def test_edit_rejects_cross_family_output_before_opening():
    opened = []

    def fake_open(*args, **kwargs):
        opened.append(args)
        return FakeWriterComposer()

    with _Monkey(api, "open_document", fake_open):
        with pytest.raises(ValueError, match="same document family"):
            api.edit("report.docx", output="report.pptx", patches=[])

    assert opened == []


def test_edit_refuses_existing_output_without_overwrite_before_opening(
    tmp_path: Path,
):
    source = tmp_path / "source.docx"
    source.write_bytes(b"source")
    output = tmp_path / "approved.docx"
    output.write_bytes(b"approved")
    opened = []

    def fake_open(*args, **kwargs):
        opened.append(args)
        return FakeWriterComposer()

    with _Monkey(api, "open_document", fake_open):
        with pytest.raises(FileExistsError, match="Output already exists"):
            api.edit(source, output=output, patches=[])

    assert output.read_bytes() == b"approved"
    assert opened == []


def test_edit_attached_refuses_existing_output_before_attach_or_mutation(
    tmp_path: Path,
):
    output = tmp_path / "approved.docx"
    output.write_bytes(b"approved")
    attach_calls = []

    def fake_attach(kind=None):
        attach_calls.append(kind)
        return FakeWriterComposer()

    with _Monkey(api, "attach_active", fake_attach):
        with pytest.raises(FileExistsError, match="Output already exists"):
            api.edit(
                output=output,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert attach_calls == []
    assert output.read_bytes() == b"approved"


def test_edit_attached_rejects_cross_family_output_before_mutation(tmp_path: Path):
    composer = FakeWriterComposer()
    output = tmp_path / "wrong-family.pptx"

    with _Monkey(api, "attach_active", lambda kind=None: composer):
        with pytest.raises(ValueError, match="same document family"):
            api.edit(
                output=output,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert composer.applied == []
    assert composer.saved is False
    assert composer.closed is True


def test_edit_attached_rejects_unsupported_save_copy_before_mutation(
    tmp_path: Path,
):
    class NoCopyComposer(FakeWriterComposer):
        def supports_attached_save_copy(self):
            return False

    composer = NoCopyComposer()

    with _Monkey(api, "attach_active", lambda kind=None: composer):
        with pytest.raises(RuntimeError, match="non-rebinding copy primitive"):
            api.edit(
                output=tmp_path / "copy.docx",
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert composer.applied == []
    assert composer.closed is True


def test_edit_attached_retains_stage_when_failed_copy_rebinds_live_document(
    tmp_path: Path,
):
    class ReboundComposer(FakeWriterComposer):
        bound_path = None
        staged_path = None

        def supports_attached_save_copy(self):
            return True

        def save_copy(self, path):
            self.staged_path = Path(path).resolve()
            with zipfile.ZipFile(self.staged_path, "w") as package:
                package.writestr("[Content_Types].xml", "<Types />")
                package.writestr("word/document.xml", "<document />")
            self.bound_path = self.staged_path
            raise RuntimeError(
                f"live document remains bound to recovery path {self.staged_path}"
            )

        def is_bound_to(self, path):
            return self.bound_path == Path(path).resolve()

    composer = ReboundComposer()

    with _Monkey(api, "attach_active", lambda kind=None: composer):
        with pytest.raises(RuntimeError, match="recovery path"):
            api.edit(
                output=tmp_path / "copy.docx",
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert composer.staged_path.is_file()
    assert not (tmp_path / "copy.docx").exists()


def test_edit_saves_to_destination_local_stage_before_publish(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"source-placeholder")
    output = tmp_path / "nested" / "edited.docx"
    composer = PackageWriterComposer()

    with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
        result = api.edit(
            source,
            output=output,
            patches=[{"target": "paragraph:1", "font": {"bold": True}}],
        )

    staged = composer.save_calls[0]
    assert staged.parent == output.parent.resolve()
    assert staged.suffix == ".docx"
    assert staged != output.resolve()
    assert not staged.exists()
    assert Path(result["saved_path"]) == output.resolve()
    assert output.is_file()


def test_edit_overwrite_save_failure_preserves_existing_output(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"source-placeholder")
    output = tmp_path / "approved.docx"
    output.write_bytes(b"approved-original")
    composer = PackageWriterComposer(fail=True)

    with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
        with pytest.raises(OSError, match="simulated save failure"):
            api.edit(
                source,
                output=output,
                overwrite=True,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert output.read_bytes() == b"approved-original"


def test_edit_corrupt_stage_does_not_replace_existing_output(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"source-placeholder")
    output = tmp_path / "approved.docx"
    output.write_bytes(b"approved-original")
    composer = PackageWriterComposer(corrupt=True)

    with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
        with pytest.raises(RuntimeError, match="STAGED|ZIP|small"):
            api.edit(
                source,
                output=output,
                overwrite=True,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert output.read_bytes() == b"approved-original"


@pytest.mark.parametrize("export_name", ["result.docx", "result"])
def test_edit_rejects_non_pdf_export_target_before_opening(
    tmp_path: Path, export_name: str
):
    opened = []

    with _Monkey(api, "open_document", lambda *args, **kwargs: opened.append(args)):
        with pytest.raises(ValueError, match="export_pdf.*\\.pdf"):
            api.edit(
                tmp_path / "source.docx",
                patches=[],
                export_pdf=tmp_path / export_name,
            )

    assert opened == []


def test_edit_refuses_existing_pdf_before_opening_or_mutation(tmp_path: Path):
    approved = tmp_path / "approved.pdf"
    approved.write_bytes(b"APPROVED-PDF")
    opened = []

    with _Monkey(api, "open_document", lambda *args, **kwargs: opened.append(args)):
        with pytest.raises(FileExistsError, match="Output already exists"):
            api.edit(
                tmp_path / "source.docx",
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
                export_pdf=approved,
            )

    assert approved.read_bytes() == b"APPROVED-PDF"
    assert opened == []


def test_edit_stages_both_outputs_before_publish_when_pdf_export_fails(
    tmp_path: Path,
):
    source = tmp_path / "source.docx"
    source.write_bytes(b"SOURCE")
    output = tmp_path / "approved.docx"
    pdf = tmp_path / "approved.pdf"
    output.write_bytes(b"APPROVED-DOCX")
    pdf.write_bytes(b"APPROVED-PDF")
    composer = PackageAndPdfWriterComposer(pdf_failure=True)

    with _Monkey(api, "_com_available", lambda: True):
        with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
            with pytest.raises(OSError, match="PDF export failure"):
                api.edit(
                    source,
                    output=output,
                    export_pdf=pdf,
                    overwrite=True,
                    patches=[{"target": "paragraph:1", "font": {"bold": True}}],
                )

    assert output.read_bytes() == b"APPROVED-DOCX"
    assert pdf.read_bytes() == b"APPROVED-PDF"
    assert composer.save_calls[0] != output.resolve()
    assert composer.export_calls[0] != pdf.resolve()
    assert not composer.save_calls[0].exists()
    assert not composer.export_calls[0].exists()


def test_edit_publishes_valid_document_and_pdf_as_one_group(tmp_path: Path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"SOURCE")
    output = tmp_path / "edited.docx"
    pdf = tmp_path / "edited.pdf"
    composer = PackageAndPdfWriterComposer()

    with _Monkey(api, "_com_available", lambda: True):
        with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
            result = api.edit(
                source,
                output=output,
                export_pdf=pdf,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert Path(result["saved_path"]) == output.resolve()
    assert Path(result["pdf_path"]) == pdf.resolve()
    assert output.is_file()
    assert pdf.is_file()


def test_file_edit_closes_composer_before_publish_and_stage_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "source.docx"
    source.write_bytes(b"SOURCE")
    output = tmp_path / "edited.docx"
    composer = PackageWriterComposer()
    events: list[str] = []
    real_publish = api.publish_artifact
    real_unlink = Path.unlink

    def close(save_changes=False):
        events.append("close")
        composer.closed = True

    def observed_publish(*args, **kwargs):
        assert composer.closed
        events.append("publish")
        return real_publish(*args, **kwargs)

    def windows_unlink(path, *args, **kwargs):
        if path in composer.save_calls:
            assert composer.closed
            events.append("cleanup")
        return real_unlink(path, *args, **kwargs)

    composer.close = close
    monkeypatch.setattr(api, "publish_artifact", observed_publish)
    monkeypatch.setattr(Path, "unlink", windows_unlink)

    with _Monkey(api, "_com_available", lambda: True):
        with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
            result = api.edit(
                source,
                output=output,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert result["saved"] is True
    assert events == ["close", "publish", "cleanup"]


def test_file_edit_closes_composer_before_group_publish_and_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "source.docx"
    source.write_bytes(b"SOURCE")
    output = tmp_path / "edited.docx"
    pdf = tmp_path / "edited.pdf"
    composer = PackageAndPdfWriterComposer()
    events: list[str] = []
    real_publish_group = api.publish_artifact_group
    real_unlink = Path.unlink

    def close(save_changes=False):
        events.append("close")
        composer.closed = True

    def observed_publish_group(*args, **kwargs):
        assert composer.closed
        events.append("publish")
        return real_publish_group(*args, **kwargs)

    def windows_unlink(path, *args, **kwargs):
        if path in [*composer.save_calls, *composer.export_calls]:
            assert composer.closed
            events.append("cleanup")
        return real_unlink(path, *args, **kwargs)

    composer.close = close
    monkeypatch.setattr(api, "publish_artifact_group", observed_publish_group)
    monkeypatch.setattr(Path, "unlink", windows_unlink)

    with _Monkey(api, "_com_available", lambda: True):
        with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
            result = api.edit(
                source,
                output=output,
                export_pdf=pdf,
                patches=[{"target": "paragraph:1", "font": {"bold": True}}],
            )

    assert result["saved"] is True
    assert result["pdf_path"] == str(pdf.resolve())
    assert events == ["close", "publish", "cleanup", "cleanup"]


def test_file_edit_close_failure_does_not_publish_and_retains_recovery_stage(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "source.docx"
    source.write_bytes(b"SOURCE")
    output = tmp_path / "edited.docx"
    composer = PackageWriterComposer()
    publish_calls = []

    def fail_close(save_changes=False):
        raise RuntimeError("simulated COM close failure")

    composer.close = fail_close
    monkeypatch.setattr(
        api, "publish_artifact", lambda *args, **kwargs: publish_calls.append(args)
    )

    with _Monkey(api, "_com_available", lambda: True):
        with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
            with pytest.raises(RuntimeError, match="recovery artifacts retained"):
                api.edit(
                    source,
                    output=output,
                    patches=[{"target": "paragraph:1", "font": {"bold": True}}],
                )

    assert publish_calls == []
    assert not output.exists()
    assert len(composer.save_calls) == 1
    assert composer.save_calls[0].is_file()


def test_edit_rejects_export_pdf_on_macos_before_bridge_or_mutation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    bridge_calls = []
    monkeypatch.setattr(api, "_com_available", lambda: False)
    from skills.WPSComposer.scripts.macos_probe import inspection
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)
    monkeypatch.setattr(
        inspection, "edit_macos", lambda *args, **kwargs: bridge_calls.append(args)
    )

    with pytest.raises(RuntimeError, match="export_pdf.*unsupported.*macOS"):
        api.edit(
            tmp_path / "source.pptx",
            export_pdf=tmp_path / "out.pdf",
            patches=[{"target": "slide:1", "fill": {"color": "#ffffff"}}],
        )

    assert bridge_calls == []


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


def test_edit_best_effort_raise_on_error_still_does_not_save():
    composer = FakeWriterComposer()

    with _Monkey(api, "open_document", lambda *args, **kwargs: composer):
        with pytest.raises(PatchError):
            api.edit(
                "report.docx",
                patches=[{"target": "boom"}],
                output="out.docx",
                atomic=False,
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
    ("slide:1/shape:@id=7/paragraph:2", "slide", "paragraph"),
    ("slide:1/shape:@id=7/paragraph:2/run:4", "slide", "run"),
    ("slide:1/shape:@id=7/table/cell:2,3", "slide", "table_cell"),
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
    # para, one None for the row end-of-mark (counted by COM Paragraphs but
    # absent from the XML), trailing body para. Missing w14:paraId -> None.
    assert _extract_paraids(_SAMPLE_DOCUMENT_XML) == [
        "0A1B2C3D", None, "FEDCBA98", None, "11223344",
    ]


def test_extract_paraids_prunes_textbox_story():
    xml = (
        b'<w:document xmlns:w="http://schemas.openxmlformats.org/'
        b'wordprocessingml/2006/main" '
        b'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        b'<w:body>'
        b'<w:p w14:paraId="0A1B2C3D"><w:r><w:drawing><w:txbxContent>'
        b'<w:p w14:paraId="BADBADBA"><w:r><w:t>textbox story</w:t></w:r></w:p>'
        b'</w:txbxContent></w:drawing></w:r></w:p>'
        b'<w:p w14:paraId="11223344"><w:r><w:t>body</w:t></w:r></w:p>'
        b'</w:body></w:document>'
    )
    # The text-frame story paragraph is not in Document.Paragraphs.
    assert _extract_paraids(xml) == ["0A1B2C3D", "11223344"]


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
        "0A1B2C3D", None, "FEDCBA98", None, "11223344",
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
    # Stable ids are rewritten to positional ones: replay targets another
    # document, where @paraId cannot resolve.
    assert targets == ["paragraph:1", "table:1/cell:1,1"]
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
                {"id": "slide:1/shape:@id=7", "index": 1, "font": {"size": 18}},
                {"id": "slide:1/shape:2", "index": 2, "fill": {"color": "#FF0000"}},
            ]},
        ],
    }
    patches = snapshot_to_patches(snapshot)
    assert {p["target"] for p in patches} == {
        "slide:1/shape:1", "slide:1/shape:2",
    }


def test_snapshot_to_patches_preserves_stable_shape_nested_scope_and_indices():
    snapshot = {
        "kind": "slide",
        "slides": [{"id": "slide:1", "shapes": [{
            "id": "slide:1/shape:@id=7",
            "index": 3,
            "fill": {"color": "#111111"},
            "paragraphs": [{
                "id": "slide:1/shape:@id=7/paragraph:2",
                "index": 2,
                "paragraph": {"alignment": 2},
                "runs": [{
                    "id": "slide:1/shape:@id=7/paragraph:2/run:4",
                    "index": 4,
                    "font": {"bold": True},
                }],
            }],
            "table": {"cells": [{
                "id": "slide:1/shape:@id=7/table/cell:2,3",
                "row": 2,
                "column": 3,
                "font": {"italic": True},
            }]},
        }]}],
    }

    assert snapshot_to_patches(snapshot) == [
        {"target": "slide:1/shape:3", "fill": {"color": "#111111"}},
        {"target": "slide:1/shape:3/paragraph:2",
         "paragraph": {"alignment": 2}},
        {"target": "slide:1/shape:3/paragraph:2/run:4",
         "font": {"bold": True}},
        {"target": "slide:1/shape:3/table/cell:2,3",
         "font": {"italic": True}},
    ]


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


def test_apply_ops_structural_rejection_is_atomic_failure():
    composer = FakeWriterComposer()

    with pytest.raises(PatchError) as exc_info:
        apply_ops(composer, [
            {"op": "remove", "target": "paragraph:405"},
        ])

    report = exc_info.value.reports[0]
    assert report["ok"] is False
    assert report["rejected"] == ["remove"]


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


def test_macos_atomic_mixed_ops_are_rejected_before_bridge_call(monkeypatch):
    from skills.WPSComposer.scripts.macos_probe import inspection

    calls = []
    monkeypatch.setattr(api, "_com_available", lambda: False)
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)
    monkeypatch.setattr(
        inspection,
        "edit_macos",
        lambda *args, **kwargs: calls.append((args, kwargs)),
    )

    result = api.edit(
        "deck.pptx",
        output="revised.pptx",
        ops=[
            {"op": "set", "target": "slide:1", "name": "Updated"},
            {"op": "remove", "target": "slide:2"},
        ],
        atomic=True,
    )

    assert result["ok"] is False
    assert result["saved"] is False
    assert result["errors"][0]["error"]["code"] == "unsupported_operation"
    assert calls == []


def test_macos_route_forwards_atomic_error_and_overwrite_semantics(monkeypatch):
    from skills.WPSComposer.scripts.macos_probe import inspection

    observed = {}
    monkeypatch.setattr(api, "_com_available", lambda: False)
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)

    def fake_edit(source, patches, output=None, **kwargs):
        observed.update(kwargs)
        return {
            "path": output,
            "saved": True,
            "patches": [{"target": "slide:1", "ok": True}],
        }

    monkeypatch.setattr(inspection, "edit_macos", fake_edit)

    result = api.edit(
        "deck.pptx",
        output="revised.pptx",
        patches=[{"target": "slide:1", "name": "Updated"}],
        atomic=False,
        raise_on_error=True,
        overwrite=True,
    )

    assert observed == {
        "atomic": False,
        "raise_on_error": True,
        "overwrite": True,
    }
    assert result["saved"] is True


def test_macos_best_effort_reports_dropped_structural_ops_as_errors(monkeypatch):
    from skills.WPSComposer.scripts.macos_probe import inspection

    monkeypatch.setattr(api, "_com_available", lambda: False)
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)
    monkeypatch.setattr(
        inspection,
        "edit_macos",
        lambda *args, **kwargs: {
            "path": "revised.pptx",
            "saved": True,
            "patches": [{"target": "slide:1", "ok": True}],
        },
    )

    result = api.edit(
        "deck.pptx",
        output="revised.pptx",
        ops=[
            {"op": "set", "target": "slide:1", "name": "Updated"},
            {"op": "remove", "target": "slide:2"},
        ],
        atomic=False,
    )

    assert result["ok"] is False
    assert result["saved"] is True
    assert result["errors"][0]["op"] == "remove"
    assert result["errors"][0]["error"]["code"] == "unsupported_operation"


def test_macos_best_effort_raise_on_error_rejects_before_bridge_publish(monkeypatch):
    from skills.WPSComposer.scripts.macos_probe import inspection

    calls = []
    monkeypatch.setattr(api, "_com_available", lambda: False)
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)
    monkeypatch.setattr(
        inspection,
        "edit_macos",
        lambda *args, **kwargs: calls.append((args, kwargs)),
    )

    with pytest.raises(PatchError) as caught:
        api.edit(
            "deck.pptx",
            output="revised.pptx",
            ops=[
                {"op": "set", "target": "slide:1", "name": "Updated"},
                {"op": "remove", "target": "slide:2"},
            ],
            atomic=False,
            raise_on_error=True,
        )

    assert caught.value.errors[0]["error"]["code"] == "unsupported_operation"
    assert calls == []


def test_macos_best_effort_structural_only_returns_error_report(monkeypatch):
    from skills.WPSComposer.scripts.macos_probe import inspection

    monkeypatch.setattr(api, "_com_available", lambda: False)
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)

    result = api.edit(
        "deck.pptx",
        ops=[{"op": "remove", "target": "slide:2"}],
        atomic=False,
    )

    assert result["ok"] is False
    assert result["saved"] is False
    assert result["ops"][0]["ok"] is False
    assert result["errors"] == result["ops"]


def test_macos_public_edit_rejects_legacy_presentation_before_com(monkeypatch):
    from skills.WPSComposer.scripts.macos_probe import inspection

    monkeypatch.setattr(api, "_com_available", lambda: False)
    monkeypatch.setattr(inspection, "macos_inspection_available", lambda: True)

    with pytest.raises(ValueError, match="macOS editing supports only '.pptx'"):
        api.edit(
            "legacy.pptm",
            patches=[{"target": "slide:1", "name": "Updated"}],
        )
