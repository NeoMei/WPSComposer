"""Checklist A: error-classification heuristic parity.
Checklist B: atomic no-save guarantee."""
import os
from pathlib import Path

from skills.WPSComposer import edit, inspect

FIX = Path(__file__).parent
OUT = FIX / "out"
OUT.mkdir(exist_ok=True)


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


def a_writer():
    r = edit(str(FIX / "sample.docx"), output=str(OUT / "a1.docx"),
             patches=[{"target": "bogus:1", "font": {"size": 20}}])
    assert r["ok"] is False, r
    kinds = {rep.get("error", {}).get("code") for rep in r["patches"]}
    assert "invalid_target" in kinds, f"expected invalid_target, got {kinds}: {r['patches']}"


def a_sheet():
    r = edit(str(FIX / "sample.xlsx"), output=str(OUT / "a2.xlsx"),
             patches=[{"target": "bogus:1", "fill": {"color": "#FF0000"}}])
    assert r["ok"] is False, r
    kinds = {rep.get("error", {}).get("code") for rep in r["patches"]}
    assert "invalid_target" in kinds, f"expected invalid_target, got {kinds}: {r['patches']}"


def a_slide():
    r = edit(str(FIX / "sample.pptx"), output=str(OUT / "a3.pptx"),
             patches=[{"target": "bogus:1", "fill": {"color": "#FF0000"}}])
    assert r["ok"] is False, r
    kinds = {rep.get("error", {}).get("code") for rep in r["patches"]}
    assert "invalid_target" in kinds, f"expected invalid_target, got {kinds}: {r['patches']}"


def b_atomic_no_save():
    out = OUT / "b_atomic.docx"
    if out.exists():
        out.unlink()
    r = edit(str(FIX / "sample.docx"), output=str(out),
             patches=[{"target": "paragraph:1", "font": {"size": 20}},
                      {"target": "paragraph:99999"}])
    assert r["ok"] is False, r
    assert r["saved"] is False, r
    assert r["saved_path"] is None, r
    assert not out.exists(), "atomic failure must not write the file"


def b_happy_path_saves():
    out = OUT / "b_happy.docx"
    if out.exists():
        out.unlink()
    r = edit(str(FIX / "sample.docx"), output=str(out),
             patches=[{"target": "paragraph:1", "font": {"size": 20}}])
    assert r["ok"] is True, r
    assert r["saved_path"].endswith("b_happy.docx"), r
    assert out.exists()


def b_source_unchanged():
    before = (FIX / "sample.docx").read_bytes()
    r = edit(str(FIX / "sample.docx"), output=str(OUT / "b_src.docx"),
             patches=[{"target": "paragraph:1", "font": {"size": 20}},
                      {"target": "paragraph:99999"}])
    assert r["ok"] is False
    after = (FIX / "sample.docx").read_bytes()
    assert before == after, "source document changed after failing atomic edit"


def main():
    results = [
        check("A.writer invalid_target", a_writer),
        check("A.sheet invalid_target", a_sheet),
        check("A.slide invalid_target", a_slide),
        check("B.atomic no-save", b_atomic_no_save),
        check("B.happy path saves", b_happy_path_saves),
        check("B.source unchanged", b_source_unchanged),
    ]
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
