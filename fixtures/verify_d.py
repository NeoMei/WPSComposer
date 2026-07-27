"""Checklist D: stable-ID readback + resolution (@paraId / @id / @name)."""
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


def d_writer_paraid_readback():
    snap = inspect(str(FIX / "sample.docx"))
    paras = snap["paragraphs"]
    assert paras, "no paragraphs in snapshot"
    with_para = [p for p in paras if p.get("para_id")]
    assert with_para, (
        "no para_ids emitted — count mismatch or missing w14:paraId; "
        f"paragraph keys: {sorted(paras[0])}, count={len(paras)}"
    )
    print(f"  ({len(with_para)}/{len(paras)} paragraphs have para_id)")


def d_writer_paraid_stable_after_structural_change():
    snap = inspect(str(FIX / "sample.docx"))
    with_para = [
        p for p in snap["paragraphs"]
        if p.get("para_id") and p.get("text") == "Sample Title"
    ]
    assert with_para, "Sample Title paragraph not found"
    pid = with_para[0]["id"]
    out = OUT / "d_writer.docx"
    # batch 1: insert a paragraph at start (positional shift), save
    r = edit(str(FIX / "sample.docx"), output=str(out),
             ops=[{"op": "insert", "type": "paragraph",
                   "props": {"text": "SHIFT"}, "position": "start"}])
    assert r["ok"], r
    # batch 2 (fresh open): patch by stable id must hit the ORIGINAL paragraph
    r = edit(str(out), output=str(out),
             patches=[{"target": pid, "font": {"bold": True}}])
    assert r["ok"], r
    snap2 = inspect(str(out))
    hit = [p for p in snap2["paragraphs"] if p.get("id") == pid]
    assert hit, f"{pid} missing after insert"
    assert hit[0]["text"] == "Sample Title", (
        f"stable id now points at different paragraph: {hit[0]['text']!r}"
    )
    assert hit[0].get("font", {}).get("bold") in (True, -1, 1), f"bold not applied: {hit[0]}"


def d_slide_id_name():
    snap = inspect(str(FIX / "sample.pptx"))
    slides = snap["slides"]
    assert slides, "no slides"
    shapes = slides[0].get("shapes", [])
    assert shapes, "no shapes on slide 1"
    sid = next((s.get("shape_id") for s in shapes if s.get("shape_id")), None)
    assert sid is not None, f"no stable shape id emitted: {shapes[0]}"
    out = OUT / "d_slide.pptx"
    r = edit(str(FIX / "sample.pptx"), output=str(out),
             patches=[{"target": f"slide:1/shape:@id={sid}",
                       "fill": {"color": "#FF0000"}}])
    assert r["ok"], r
    # name form
    name = shapes[0].get("name")
    assert name, "shape has no name"
    r = edit(str(FIX / "sample.pptx"), output=str(out),
             patches=[{"target": f"slide:1/shape:@name={name}",
                       "fill": {"color": "#00FF00"}}])
    assert r["ok"], r


def d_sheet_id_name():
    snap = inspect(str(FIX / "sample.xlsx"))
    sheets = snap["sheets"]
    assert sheets, "no sheets"
    shapes = sheets[0].get("shapes", [])
    assert shapes, "no shapes on sheet 1 (chart missing?)"
    sid = next((s.get("shape_id") for s in shapes if s.get("shape_id")), None)
    assert sid is not None, f"no stable shape id emitted: {shapes[0]}"
    out = OUT / "d_sheet.xlsx"
    r = edit(str(FIX / "sample.xlsx"), output=str(out),
             patches=[{"target": f"sheet:1/shape:@id={sid}",
                       "line": {"color": "#FF0000"}}])
    assert r["ok"], r
    name = shapes[0].get("name")
    assert name, "shape has no name"
    r = edit(str(FIX / "sample.xlsx"), output=str(out),
             patches=[{"target": f"sheet:1/shape:@name={name}",
                       "line": {"color": "#00FF00"}}])
    assert r["ok"], r


def main():
    results = [
        check("D.writer paraId readback", d_writer_paraid_readback),
        check("D.writer paraId stable after structural change",
              d_writer_paraid_stable_after_structural_change),
        check("D.slide @id/@name", d_slide_id_name),
        check("D.sheet @id/@name", d_sheet_id_name),
    ]
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
