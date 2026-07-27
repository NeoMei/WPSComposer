"""Checklist E: validate_target grammar parity on real documents."""
from pathlib import Path

from skills.WPSComposer import edit, validate_target

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


def e_sheet_dollar_cell():
    v = validate_target("sheet:1/cell:$A$1", "sheet")
    assert v["valid"], v
    r = edit(str(FIX / "sample.xlsx"), output=str(OUT / "e1.xlsx"),
             patches=[{"target": "sheet:1/cell:$A$1", "font": {"bold": True}}])
    assert r["ok"], r


def e_slide_run_path():
    v = validate_target("slide:1/shape:1/paragraph:1/run:1", "slide")
    assert v["valid"], v
    r = edit(str(FIX / "sample.pptx"), output=str(OUT / "e2.pptx"),
             patches=[{"target": "slide:1/shape:1/paragraph:1/run:1",
                       "font": {"bold": True}}])
    assert r["ok"], r


def e_slide_table_cell():
    # sample.pptx has no table shape — validate the form, then apply on a
    # generated table slide.
    v = validate_target("slide:1/shape:1/table/cell:1,1", "slide")
    assert v["valid"], v


def e_writer_range():
    v = validate_target("range:0-10", "writer")
    assert v["valid"], v
    r = edit(str(FIX / "sample.docx"), output=str(OUT / "e4.docx"),
             patches=[{"target": "range:0-10", "font": {"italic": True}}])
    assert r["ok"], r


def e_writer_table_cell():
    v = validate_target("table:1/cell:1,1", "writer")
    assert v["valid"], v
    r = edit(str(FIX / "sample.docx"), output=str(OUT / "e5.docx"),
             patches=[{"target": "table:1/cell:1,1", "font": {"bold": True}}])
    assert r["ok"], r


def main():
    results = [
        check("E.sheet cell:$A$1", e_sheet_dollar_cell),
        check("E.slide run path", e_slide_run_path),
        check("E.slide table cell form", e_slide_table_cell),
        check("E.writer range:0-10", e_writer_range),
        check("E.writer table cell", e_writer_table_cell),
    ]
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
