"""Checklist C: attach-active caveat — atomic failure skips save_current(),
but the live window may still show partially-applied formatting."""
import shutil
from pathlib import Path

from skills.WPSComposer import edit
from skills.WPSComposer.scripts._dispatch import _dispatch

FIX = Path(__file__).parent
WORK = FIX / "out" / "c_live.docx"
shutil.copy2(FIX / "sample.docx", WORK)


def main():
    results = []

    # open the doc visibly in a dedicated WPS instance
    app = _dispatch(("KWps.Application", "Wps.Application", "Word.Application"))
    app.Visible = -1
    doc = app.Documents.Open(str(WORK.resolve()))
    before = WORK.read_bytes()

    # failing atomic batch: first patch applies (visible in the live doc),
    # second fails -> no save
    r = edit(None, kind="writer", patches=[
        {"target": "paragraph:2", "font": {"size": 20}},
        {"target": "paragraph:99999"},
    ])
    results.append(("no-save on atomic failure",
                    r["ok"] is False and r["saved"] is False))
    results.append(("disk unchanged", WORK.read_bytes() == before))

    # caveat half: the live document DID receive the first patch
    size_now = doc.Paragraphs(2).Range.Font.Size
    results.append(("live window shows partial formatting (documented caveat)",
                    float(size_now) == 20.0))

    # happy path through attach: saves in place
    r = edit(None, kind="writer",
             patches=[{"target": "paragraph:2", "font": {"size": 22}}])
    results.append(("attach happy path saves", r["ok"] is True and r["saved"] is True))

    doc.Close(False)
    app.Quit()

    for name, ok in results:
        print(("PASS" if ok else "FAIL"), name)
    print(f"{sum(1 for _, ok in results if ok)}/{len(results)} passed")
    return 0 if all(ok for _, ok in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
