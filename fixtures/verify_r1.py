"""Completeness round 1: coverage gaps found in the audit.
- Writer: move table, move shape, clone shape (clipboard matrix)
- Slide: table cell target actually applies (slide:N/shape:N/table/cell:R,C)
- Writer paraId alignment on more document shapes (multi-table, no-textbox)
"""
from pathlib import Path

from skills.WPSComposer import edit, generate, inspect
from skills.WPSComposer.scripts.writer import read_paraids_from_docx
from skills.WPSComposer.scripts.slide import SlideComposer

FIX = Path(__file__).parent
OUT = FIX / "out"


def check(name, fn):
    try:
        fn()
        print(f"PASS {name}")
        return True
    except AssertionError as exc:
        print(f"FAIL {name}: {exc}")
        return False
    except Exception as exc:
        print(f"ERROR {name}: {type(exc).__name__}: {exc}")
        return False


def r1_writer_move_table():
    out = OUT / "r1_w1.docx"
    before = inspect(str(FIX / "sample.docx"))
    texts = [p["text"] for p in before["paragraphs"]]
    closing_pos = texts.index("Closing paragraph for the sample.")
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "move", "target": "table:1", "to": "end"}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["tables"] == 1, snap["counts"]
    # table now after the closing paragraph
    texts2 = [p["text"] for p in snap["paragraphs"]]
    assert "Closing paragraph for the sample." in texts2


def r1_writer_move_shape():
    out = OUT / "r1_w2.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "move", "target": "shape:1", "to": "end"}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["shapes"] == 1, snap["counts"]


def r1_writer_clone_shape():
    out = OUT / "r1_w3.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "clone", "target": "shape:1", "to": "end"}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["shapes"] == 2, snap["counts"]


def r1_slide_table_cell_apply():
    # build a pptx with a real table shape via the composer API
    md = FIX / "_tbl.md"
    md.write_text("# T\n\nbody\n", encoding="utf-8")
    base = OUT / "r1_tbl_base.pptx"
    base.unlink(missing_ok=True)
    generate(str(md), format="pptx", output=str(base))
    md.unlink()
    with SlideComposer.open_document(str(base)) as comp:
        comp.add_table(1, 2, 2, 50, 150, 400, 150,
                       data=[["h1", "h2"], ["a", "b"]])
        comp.save(str(base))
    snap = inspect(str(base))
    tbl_pos = None
    for si, slide in enumerate(snap["slides"], start=1):
        for sh in slide.get("shapes", []):
            if sh.get("has_table") or sh.get("table"):
                tbl_pos = (si, sh["index"])
    assert tbl_pos, "no table shape found after add_table"
    out = OUT / "r1_s1.pptx"
    r = edit(str(base), output=str(out), patches=[
        {"target": f"slide:{tbl_pos[0]}/shape:{tbl_pos[1]}/table/cell:1,1",
         "font": {"bold": True}}])
    assert r["ok"], r


def r1_paraid_alignment_more_docs():
    cases = {
        "multi_table": "# H\n\npara one\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n"
                       "mid\n\n| C | D | E |\n|---|---|---|\n| 3 | 4 | 5 |\n| 6 | 7 | 8 |\n\nend\n",
        "no_textbox_simple": "# Only paragraphs\n\nalpha\n\nbeta\n\ngamma\n",
    }
    import win32com.client as wc
    for name, content in cases.items():
        md = FIX / f"_pa_{name}.md"
        md.write_text(content, encoding="utf-8")
        docx = OUT / f"pa_{name}.docx"
        docx.unlink(missing_ok=True)
        generate(str(md), format="docx", output=str(docx))
        md.unlink()
        xml_count = len(read_paraids_from_docx(docx))
        app = wc.DispatchEx("Kwps.Application")
        app.Visible = False
        doc = app.Documents.Open(str(docx.resolve()), ReadOnly=True)
        com_count = doc.Paragraphs.Count
        doc.Close(False)
        app.Quit()
        assert xml_count == com_count, (
            f"{name}: xml {xml_count} != com {com_count}")


def main():
    results = [
        check("R1.writer move table", r1_writer_move_table),
        check("R1.writer move shape", r1_writer_move_shape),
        check("R1.writer clone shape", r1_writer_clone_shape),
        check("R1.slide table cell apply", r1_slide_table_cell_apply),
        check("R1.paraid alignment (multi-table, simple)", r1_paraid_alignment_more_docs),
    ]
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
