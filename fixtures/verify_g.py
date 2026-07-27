"""Checklist G: structural verbs full matrix (insert/remove/move/clone)."""
from pathlib import Path

from skills.WPSComposer import edit, inspect

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


def g_writer_insert_paragraph_end():
    out = OUT / "g_w1.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "insert", "type": "paragraph", "props": {"text": "appended"}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert any(p["text"] == "appended" for p in snap["paragraphs"]), \
        [p["text"] for p in snap["paragraphs"][-3:]]


def g_writer_insert_heading_end():
    out = OUT / "g_w2.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "insert", "type": "heading",
         "props": {"text": "New Heading", "level": 1}}])
    assert r["ok"], r
    snap = inspect(str(out))
    hit = [p for p in snap["paragraphs"] if p["text"] == "New Heading"]
    assert hit, "heading text missing"
    style = str(hit[0].get("style", ""))
    assert "Heading" in style or "标题" in style, f"Heading style not applied: {style!r}"


def g_writer_insert_table():
    out = OUT / "g_w3.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "insert", "type": "table",
         "props": {"rows": 2, "cols": 2, "data": [["a", "b"], ["c", "d"]]}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["tables"] == 2, snap["counts"]


def g_writer_insert_image_and_textbox():
    out = OUT / "g_w4.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "insert", "type": "image",
         "props": {"path": str(FIX / "logo.png")}},
        {"op": "insert", "type": "textbox",
         "props": {"text": "note", "left": 100, "top": 300,
                   "width": 200, "height": 50}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["inline_shapes"] == 2, snap["counts"]
    assert snap["counts"]["shapes"] == 2, snap["counts"]


def g_writer_insert_page_break():
    out = OUT / "g_w5.docx"
    before = inspect(str(FIX / "sample.docx"))["counts"]["paragraphs"]
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "insert", "type": "page_break"}])
    assert r["ok"], r


def g_writer_remove_table_and_shape():
    out = OUT / "g_w6.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "remove", "target": "table:1"},
        {"op": "remove", "target": "shape:1"}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["tables"] == 0, snap["counts"]
    assert snap["counts"]["shapes"] == 0, snap["counts"]


def g_writer_remove_paragraph():
    out = OUT / "g_w7.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "remove", "target": "paragraph:3"}])
    assert r["ok"], r
    snap = inspect(str(out))
    texts = [p["text"] for p in snap["paragraphs"]]
    assert len(snap["paragraphs"]) == 29, len(snap["paragraphs"])


def g_writer_clone_paragraph():
    out = OUT / "g_w8.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "clone", "target": "paragraph:2", "to": "end"}])
    assert r["ok"], r
    snap = inspect(str(out))
    texts = [p["text"] for p in snap["paragraphs"]]
    assert texts.count("Sample Title") == 2, texts


def g_writer_move_paragraph():
    out = OUT / "g_w9.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "move", "target": "paragraph:2",
         "to": {"after": "paragraph:4"}}])
    assert r["ok"], r
    snap = inspect(str(out))
    texts = [p["text"] for p in snap["paragraphs"]]
    assert "Sample Title" in texts, texts[:8]
    # moved title now sits after the former paragraph 4
    title_pos = texts.index("Sample Title")
    assert title_pos > 1, texts[:8]


def g_writer_clone_table():
    out = OUT / "g_w10.docx"
    r = edit(str(FIX / "sample.docx"), output=str(out), ops=[
        {"op": "clone", "target": "table:1", "to": "end"}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert snap["counts"]["tables"] == 2, snap["counts"]


def g_slide_insert_slide():
    out = OUT / "g_s1.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "insert", "type": "slide", "props": {"layout": 12}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert len(snap["slides"]) == 5, len(snap["slides"])


def g_slide_insert_textbox():
    out = OUT / "g_s2.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "insert", "parent": "slide:1", "type": "textbox",
         "props": {"text": "inserted box", "left": 50, "top": 50,
                   "width": 200, "height": 50}}])
    assert r["ok"], r
    snap = inspect(str(out))
    texts = [s.get("text", "") for s in snap["slides"][0].get("shapes", [])]
    assert any("inserted box" in t for t in texts), texts


def g_slide_insert_image():
    out = OUT / "g_s3.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "insert", "parent": "slide:2", "type": "image",
         "props": {"path": str(FIX / "logo.png"), "left": 10, "top": 10}}])
    assert r["ok"], r


def g_slide_remove_shape():
    out = OUT / "g_s4.pptx"
    before = inspect(str(FIX / "sample.pptx"))
    n = len(before["slides"][0].get("shapes", []))
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "remove", "target": "slide:1/shape:1"}])
    assert r["ok"], r
    after = inspect(str(out))
    assert len(after["slides"][0].get("shapes", [])) == n - 1, (
        n, len(after["slides"][0].get("shapes", [])))


def g_slide_remove_slide():
    out = OUT / "g_s5.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "remove", "target": "slide:2"}])
    assert r["ok"], r
    assert len(inspect(str(out))["slides"]) == 3


def g_slide_move_slide():
    out = OUT / "g_s6.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "move", "target": "slide:1", "to": {"after": "slide:3"}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert len(snap["slides"]) == 4


def g_slide_move_shape_cross_slide():
    out = OUT / "g_s7.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "move", "target": "slide:1/shape:1", "to": {"slide": 3}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert len(snap["slides"][2].get("shapes", [])) >= 1


def g_slide_clone_slide():
    out = OUT / "g_s8.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "clone", "target": "slide:1", "to": {"after": "slide:2"}}])
    assert r["ok"], r
    assert len(inspect(str(out))["slides"]) == 5


def g_slide_clone_shape_same_and_cross():
    snap = inspect(str(FIX / "sample.pptx"))
    sid = snap["slides"][0]["shapes"][0].get("shape_id")
    assert sid is not None, "no stable shape id on slide 1"
    out = OUT / "g_s9.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out), ops=[
        {"op": "clone", "target": "slide:1/shape:1"},
        {"op": "clone", "target": f"slide:1/shape:@id={sid}", "to": {"slide": 2}}])
    assert r["ok"], r
    snap2 = inspect(str(out))
    assert len(snap2["slides"][1].get("shapes", [])) >= 2, \
        [len(s.get("shapes", [])) for s in snap2["slides"]]


def g_sheet_insert_row_column():
    out = OUT / "g_x1.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "insert", "parent": "sheet:1", "type": "row",
         "props": {"values": ["new", "row", "99"]}, "position": {"index": 2}},
        {"op": "insert", "parent": "sheet:1", "type": "column",
         "props": {"values": ["x", "y"]}, "position": {"index": 2}}])
    assert r["ok"], r
    snap = inspect(str(out))
    sheet = snap["sheets"][0]
    assert sheet.get("used_rows") == 5, sheet.get("used_rows")
    assert sheet.get("used_columns") == 4, sheet.get("used_columns")


def _cells(sheet):
    cells = sheet.get("cells", [])
    if isinstance(cells, dict):
        return cells
    return {c.get("address") or c.get("cell") or c.get("id"): c for c in cells}


def g_sheet_insert_sheet():
    out = OUT / "g_x2.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "insert", "type": "sheet", "props": {"name": "Summary"}}])
    assert r["ok"], r
    snap = inspect(str(out))
    names = [s.get("name") for s in snap["sheets"]]
    assert "Summary" in names, names


def g_sheet_remove_row():
    out = OUT / "g_x3.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "remove", "target": "sheet:1/cell:C1"}])  # axis=row default
    assert r["ok"], r
    snap = inspect(str(out))
    cells = _cells(snap["sheets"][0])
    a1 = cells.get("$A$1", {})
    assert a1.get("value") == "apples", f"row 1 not removed: A1={a1}"


def g_sheet_remove_column():
    out = OUT / "g_x4.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "remove", "target": "sheet:1/cell:C1", "axis": "column"}])
    assert r["ok"], r


def g_sheet_remove_chart():
    out = OUT / "g_x5.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "remove", "target": "sheet:1/chart:1"}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert len(snap["sheets"][0].get("charts", [])) == 0, snap["sheets"][0].get("charts")


def g_sheet_remove_shape():
    out = OUT / "g_x6.xlsx"
    snap = inspect(str(FIX / "sample.xlsx"))
    shapes = snap["sheets"][0].get("shapes", [])
    assert shapes, "no shapes to remove"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "remove", "target": f"sheet:1/shape:@id={shapes[0]['shape_id']}"}])
    assert r["ok"], r


def g_sheet_clone_row():
    out = OUT / "g_x7.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "clone", "target": "sheet:1/cell:A2", "to": {"index": 5}}])
    assert r["ok"], r
    snap = inspect(str(out))
    cells = _cells(snap["sheets"][0])
    assert cells.get("$A$5", {}).get("value") == "apples", cells.get("$A$5")


def g_sheet_clone_and_move_sheet():
    out = OUT / "g_x8.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "clone", "target": "sheet:1", "to": {"after": 1}},
        {"op": "move", "target": "sheet:1", "to": {"after": 2}}])
    assert r["ok"], r
    snap = inspect(str(out))
    assert len(snap["sheets"]) == 2, len(snap["sheets"])
    assert snap["sheets"][1].get("name") != snap["sheets"][0].get("name")


def g_sheet_remove_last_sheet_guarded():
    out = OUT / "g_x9.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out), ops=[
        {"op": "remove", "target": "sheet:1"}])  # only sheet -> must fail clean
    assert r["ok"] is False, r
    codes = {e.get("error", {}).get("code") for e in r["errors"]}
    assert codes & {"invalid_target", "apply_failed"}, codes
    assert not out.exists(), "failure must not save"


def _slide_uids(snap):
    uids = []
    for s in snap["slides"]:
        uid = next((sh.get("text", "").strip()
                    for sh in s.get("shapes", [])
                    if sh.get("text", "").strip().startswith("UID-")), "?")
        uids.append(uid)
    return uids


def _stamp_ops(n):
    # sample.pptx reuses "Deck Title" on slides 1-2; stamp each slide with a
    # unique textbox so move/clone landing positions are unambiguous.
    return [{"op": "insert", "parent": f"slide:{i}", "type": "textbox",
             "props": {"text": f"UID-{chr(64 + i)}", "left": 500, "top": 400,
                       "width": 100, "height": 30}}
            for i in range(1, n + 1)]


def g_slide_move_after_last_anchor():
    # Review-round regression: source < anchor, anchor is LAST slide.
    out = OUT / "g_s10.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out),
             ops=_stamp_ops(4) + [
                 {"op": "move", "target": "slide:1", "to": {"after": "slide:4"}}])
    assert r["ok"], r
    uids = _slide_uids(inspect(str(out)))
    assert uids == ["UID-B", "UID-C", "UID-D", "UID-A"], \
        f"slide 1 must land last: {uids}"


def g_slide_move_before_anchor_source_before():
    # Review-round regression: source < anchor with 'before'.
    out = OUT / "g_s11.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out),
             ops=_stamp_ops(4) + [
                 {"op": "move", "target": "slide:1", "to": {"before": "slide:3"}}])
    assert r["ok"], r
    uids = _slide_uids(inspect(str(out)))
    assert uids == ["UID-B", "UID-A", "UID-C", "UID-D"], \
        f"slide 1 must land before former slide 3: {uids}"


def g_slide_clone_to_end():
    # Review-round regression: clone with to='end' must not crash and lands last.
    out = OUT / "g_s12.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out),
             ops=_stamp_ops(4) + [
                 {"op": "clone", "target": "slide:1", "to": "end"}])
    assert r["ok"], r
    uids = _slide_uids(inspect(str(out)))
    assert uids == ["UID-A", "UID-B", "UID-C", "UID-D", "UID-A"], \
        f"clone must land last: {uids}"


def main():
    results = [
        check("G.writer insert paragraph@end", g_writer_insert_paragraph_end),
        check("G.writer insert heading@end (style!)", g_writer_insert_heading_end),
        check("G.writer insert table", g_writer_insert_table),
        check("G.writer insert image+textbox", g_writer_insert_image_and_textbox),
        check("G.writer insert page_break", g_writer_insert_page_break),
        check("G.writer remove table+shape", g_writer_remove_table_and_shape),
        check("G.writer remove paragraph", g_writer_remove_paragraph),
        check("G.writer clone paragraph", g_writer_clone_paragraph),
        check("G.writer move paragraph", g_writer_move_paragraph),
        check("G.writer clone table", g_writer_clone_table),
        check("G.slide insert slide", g_slide_insert_slide),
        check("G.slide insert textbox", g_slide_insert_textbox),
        check("G.slide insert image", g_slide_insert_image),
        check("G.slide remove shape", g_slide_remove_shape),
        check("G.slide remove slide", g_slide_remove_slide),
        check("G.slide move slide", g_slide_move_slide),
        check("G.slide move shape cross-slide", g_slide_move_shape_cross_slide),
        check("G.slide clone slide", g_slide_clone_slide),
        check("G.slide clone shape same+cross", g_slide_clone_shape_same_and_cross),
        check("G.sheet insert row+column", g_sheet_insert_row_column),
        check("G.sheet insert sheet", g_sheet_insert_sheet),
        check("G.sheet remove row", g_sheet_remove_row),
        check("G.sheet remove column", g_sheet_remove_column),
        check("G.sheet remove chart", g_sheet_remove_chart),
        check("G.sheet remove shape", g_sheet_remove_shape),
        check("G.sheet clone row", g_sheet_clone_row),
        check("G.sheet clone+move sheet", g_sheet_clone_and_move_sheet),
        check("G.sheet last-sheet removal guarded", g_sheet_remove_last_sheet_guarded),
        check("G.slide move after LAST anchor", g_slide_move_after_last_anchor),
        check("G.slide move before anchor (src<anchor)", g_slide_move_before_anchor_source_before),
        check("G.slide clone to='end'", g_slide_clone_to_end),
    ]
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
